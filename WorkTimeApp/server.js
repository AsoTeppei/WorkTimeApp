// ============================================================
// 作業時間 一元管理システム — API サーバー（要件定義 v2 対応）
//   ・営業グループ追加
//   ・目標時間（ProjectTargets）対応
//   ・部署マスタ管理エンドポイント追加
//   ・作業記録の論理削除 / 注番の編集・クローズ / 目標時間の更新
//   ・ログのページネーション
//
// 【必要なパッケージのインストール】
//   npm install express cors mssql
//
// 【起動方法】
//   node server.js
// ============================================================

// ------------------------------------------------------------
// stdout/stderr をブロッキングモードに切り替える
//   Task Scheduler 経由で start_server.bat 経由で起動した場合、
//   stdout はパイプ扱いになりブロックバッファリングされる。
//   その結果 console.log の出力がバッファに溜まり続け、
//   logs/server.log にほとんど書き出されない。
//   _handle.setBlocking(true) で同期書き込みに切り替えると、
//   1 行ごとに即座にファイルへ反映される。
// ------------------------------------------------------------
if (process.stdout._handle && typeof process.stdout._handle.setBlocking === 'function') {
    process.stdout._handle.setBlocking(true);
}
if (process.stderr._handle && typeof process.stderr._handle.setBlocking === 'function') {
    process.stderr._handle.setBlocking(true);
}

const express = require('express');
const cors    = require('cors');
const sql     = require('mssql');
const crypto  = require('crypto');
const path    = require('path');

const app  = express();
const PORT = 3000;

app.use(express.json());
app.use(cors());

// ============================================================
// 静的配信: index.html をブラウザに返す
//
//   http://192.168.1.8:3000/          → index.html
//   http://192.168.1.8:3000/index.html → index.html
//
// セキュリティ方針:
//   express.static(__dirname) でプロジェクト全体を公開すると
//   server.js（SQL 認証情報を平文で持つ）、.bat、.sql、logs/ 等が
//   全て読めてしまう。ここでは index.html だけをホワイトリスト方式で
//   公開する。他ファイルが必要になったら、個別に app.get() を追加する。
// ============================================================
const INDEX_HTML_PATH = path.join(__dirname, 'index.html');
app.get('/', (req, res) => res.sendFile(INDEX_HTML_PATH));
app.get('/index.html', (req, res) => res.sendFile(INDEX_HTML_PATH));

// ============================================================
// 認証用 秘密鍵（本番運用時は環境変数化を推奨）
// ※ この値を変更すると、全ユーザーが強制ログアウトになります
// ============================================================
const AUTH_SECRET = 'worktime-lan-hmac-9f8a3c2b1d4e5f6789abcdef0123456789';
// トークン有効期限: 実質無期限（10年）— ログアウトするまでPIN再入力不要
const TOKEN_TTL_MS = 10 * 365 * 24 * 60 * 60 * 1000;

// ============================================================
// PIN ハッシュ化（scrypt + salt）
// ============================================================
function hashPin(pin) {
    const salt = crypto.randomBytes(16).toString('hex');
    const hash = crypto.scryptSync(String(pin), salt, 64).toString('hex');
    return `${salt}:${hash}`;
}
function verifyPin(pin, stored) {
    if (!stored || typeof stored !== 'string' || !stored.includes(':')) return false;
    const [salt, hash] = stored.split(':');
    try {
        const test = crypto.scryptSync(String(pin), salt, 64).toString('hex');
        const a = Buffer.from(hash, 'hex');
        const b = Buffer.from(test, 'hex');
        if (a.length !== b.length) return false;
        return crypto.timingSafeEqual(a, b);
    } catch (e) { return false; }
}

// ============================================================
// 認証トークン（ステートレス HMAC）
// 形式: userId.role.tokenVersion.expiry.signature
// ============================================================
function createToken(userId, role, tokenVersion) {
    const expiry  = Date.now() + TOKEN_TTL_MS;
    const payload = `${userId}.${role}.${tokenVersion}.${expiry}`;
    const sig     = crypto.createHmac('sha256', AUTH_SECRET).update(payload).digest('hex');
    return `${payload}.${sig}`;
}
function verifyToken(token) {
    if (!token || typeof token !== 'string') return null;
    const parts = token.split('.');
    if (parts.length !== 5) return null;
    const [userId, role, tokenVersion, expiry, sig] = parts;
    const payload  = `${userId}.${role}.${tokenVersion}.${expiry}`;
    const expected = crypto.createHmac('sha256', AUTH_SECRET).update(payload).digest('hex');
    try {
        const a = Buffer.from(sig, 'hex');
        const b = Buffer.from(expected, 'hex');
        if (a.length !== b.length) return null;
        if (!crypto.timingSafeEqual(a, b)) return null;
    } catch (e) { return null; }
    if (Number(expiry) < Date.now()) return null;
    return { userId: Number(userId), role, tokenVersion: Number(tokenVersion) };
}

// ============================================================
// 認証ミドルウェア
//   - Authorization: Bearer <token> を検証
//   - DB の TokenVersion と突合して、古いセッションを無効化
//   - 成功: req.user = { userId, role } をセット
// ============================================================
async function authMiddleware(req, res, next) {
    const header = req.headers.authorization || '';
    const token  = header.startsWith('Bearer ') ? header.slice(7) : '';
    const claims = verifyToken(token);
    if (!claims) return res.status(401).json({ error: '認証が必要です', code: 'UNAUTHORIZED' });

    try {
        const pool = await getPool();
        const r = await pool.request()
            .input('UserID', sql.Int, claims.userId)
            .query('SELECT UserID, Name, Role, TokenVersion, IsActive FROM Users WHERE UserID = @UserID');
        const row = r.recordset[0];
        if (!row || !row.IsActive) {
            return res.status(401).json({ error: 'ユーザーが見つかりません', code: 'UNAUTHORIZED' });
        }
        if (Number(row.TokenVersion) !== Number(claims.tokenVersion)) {
            return res.status(401).json({ error: 'セッションが無効になりました。再ログインしてください', code: 'TOKEN_STALE' });
        }
        req.user = { userId: row.UserID, name: row.Name, role: row.Role };
        next();
    } catch (err) {
        console.error('authMiddleware エラー:', err);
        res.status(500).json({ error: 'Auth check failed', details: err.message });
    }
}

// ロール別ガード
function requireRole(...roles) {
    return (req, res, next) => {
        if (!req.user) return res.status(401).json({ error: '認証が必要です' });
        if (!roles.includes(req.user.role)) {
            return res.status(403).json({ error: 'この操作を実行する権限がありません', code: 'FORBIDDEN' });
        }
        next();
    };
}

// PIN の形式チェック（4桁の半角数字のみ）
function isValidPin(pin) {
    return typeof pin === 'string' && /^\d{4}$/.test(pin);
}

// ============================================================
// データベース接続設定
// ============================================================
const dbConfig = {
    user:     'yonekura',
    password: 'yone6066',
    server:   '192.168.1.8',
    database: 'WorkTimeDB',
    options: {
        encrypt:                false,
        trustServerCertificate: true,
        instanceName:           'SQLEXPRESS',
        // 接続のキープアライブを有効化。ファイアウォール/NAT がアイドル接続を
        // 切るのを防ぐ。20秒ごとに TCP keepalive を送る。
        enableArithAbort:       true
    },
    connectionTimeout: 15000,    // 接続確立のタイムアウト 15秒
    requestTimeout:    15000,    // 個別クエリのタイムアウト 15秒
    // ★ コネクションプール設定
    //   LAN 内 30 ユーザー前後 + 30 秒の自動同期を想定。
    //   - acquireTimeoutMillis を短くして、プールが詰まったときに長時間
    //     ユーザーを待たせない（自己回復が走るまでの最大遅延）。
    //   - idleTimeoutMillis を短くして、アイドル接続を早めに破棄し、
    //     ファイアウォール/NAT による片側切断を回避する。
    pool: {
        max: 20,
        min: 2,
        idleTimeoutMillis:    30000,    // 30秒で idle 接続を破棄
        acquireTimeoutMillis: 8000,     // 8秒で諦めて自己回復
        createTimeoutMillis:  10000,    // 接続作成は 10秒で諦める
        destroyTimeoutMillis: 5000,
        reapIntervalMillis:   1000,     // 1秒ごとに idle/expired をスキャン
        createRetryIntervalMillis: 200
    }
};

// ============================================================
// 注番自動クローズ設定
//   最終入力から AUTO_CLOSE_INACTIVE_MONTHS ヶ月以上経った進行中の注番を
//   サーバー起動時と日次で自動ロックする。
// ============================================================
const AUTO_CLOSE_INACTIVE_MONTHS = 2;
const AUTO_CLOSE_INTERVAL_MS     = 24 * 60 * 60 * 1000; // 24時間

// ============================================================
// 日本の祝日リスト（手動メンテ）
//
//   入力忘れ通知機能（GET /api/me/input-status）で使用。
//   土日は dow で判定し、祝日はここに列挙されたものを除外する。
//
//   ★ 運用メモ:
//     - 春分の日・秋分の日は前年2月の官報で翌年分が確定する。
//       未確定年は暫定値を入れてあるので、官報発表後に確認・修正する。
//     - 毎年 12 月頃に「翌年分が入っているか」を確認して追加する。
//     - 会社独自の休業日（年末年始、夏季休業など）は含まれていない。
//       必要なら COMPANY_HOLIDAYS に追加する。
// ============================================================
const JAPANESE_HOLIDAYS = new Set([
    // ---- 2025 ----
    '2025-01-01', // 元日
    '2025-01-13', // 成人の日
    '2025-02-11', // 建国記念の日
    '2025-02-23', // 天皇誕生日
    '2025-02-24', // 振替休日
    '2025-03-20', // 春分の日
    '2025-04-29', // 昭和の日
    '2025-05-03', // 憲法記念日
    '2025-05-04', // みどりの日
    '2025-05-05', // こどもの日
    '2025-05-06', // 振替休日
    '2025-07-21', // 海の日
    '2025-08-11', // 山の日
    '2025-09-15', // 敬老の日
    '2025-09-23', // 秋分の日
    '2025-10-13', // スポーツの日
    '2025-11-03', // 文化の日
    '2025-11-23', // 勤労感謝の日
    '2025-11-24', // 振替休日
    // ---- 2026 ----
    '2026-01-01', // 元日
    '2026-01-12', // 成人の日
    '2026-02-11', // 建国記念の日
    '2026-02-23', // 天皇誕生日
    '2026-03-20', // 春分の日
    '2026-04-29', // 昭和の日
    '2026-05-03', // 憲法記念日
    '2026-05-04', // みどりの日
    '2026-05-05', // こどもの日
    '2026-05-06', // 振替休日
    '2026-07-20', // 海の日
    '2026-08-11', // 山の日
    '2026-09-21', // 敬老の日
    '2026-09-22', // 国民の休日
    '2026-09-23', // 秋分の日
    '2026-10-12', // スポーツの日
    '2026-11-03', // 文化の日
    '2026-11-23', // 勤労感謝の日
    // ---- 2027 ----（暫定）
    '2027-01-01', // 元日
    '2027-01-11', // 成人の日
    '2027-02-11', // 建国記念の日
    '2027-02-23', // 天皇誕生日
    '2027-03-21', // 春分の日（暫定）
    '2027-03-22', // 振替休日
    '2027-04-29', // 昭和の日
    '2027-05-03', // 憲法記念日
    '2027-05-04', // みどりの日
    '2027-05-05', // こどもの日
    '2027-07-19', // 海の日
    '2027-08-11', // 山の日
    '2027-09-20', // 敬老の日
    '2027-09-23', // 秋分の日（暫定）
    '2027-10-11', // スポーツの日
    '2027-11-03', // 文化の日
    '2027-11-23', // 勤労感謝の日
]);

// 会社独自の休業日（必要時に追加）
//   例: '2026-12-30', '2026-12-31' など
const COMPANY_HOLIDAYS = new Set([
]);

// YYYY-MM-DD 形式の文字列を「ローカル時刻」の Date にして、
// 同じフォーマットに戻す軽量ヘルパー
function formatDateLocal(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

// 営業日判定（土日・祝日・会社休業日を除外）
function isBusinessDay(dateStr) {
    // dateStr: 'YYYY-MM-DD' (ローカル日付)
    const [y, m, d] = dateStr.split('-').map(Number);
    const dt = new Date(y, m - 1, d);
    const dow = dt.getDay();
    if (dow === 0 || dow === 6) return false;     // 土日
    if (JAPANESE_HOLIDAYS.has(dateStr)) return false;
    if (COMPANY_HOLIDAYS.has(dateStr))  return false;
    return true;
}

// 指定日（デフォルト: 今日）から遡って直近 N 営業日の日付配列を返す
// 返り値は古い順（例: ['2026-04-07', '2026-04-08', ..., '2026-04-15']）
function lastNBusinessDays(n, referenceDate = null) {
    const ref = referenceDate || new Date();
    const result = [];
    const cursor = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate());
    // 暴走防止: 最大 n*3 日まで遡る
    let safety = n * 3 + 30;
    while (result.length < n && safety-- > 0) {
        const dateStr = formatDateLocal(cursor);
        if (isBusinessDay(dateStr)) {
            result.unshift(dateStr);
        }
        cursor.setDate(cursor.getDate() - 1);
    }
    return result;
}

// ============================================================
// 共通: DBプール取得（シングルトン + 自己回復機能）
//
//   mssql 6.x+ は sql.connect(config) が内部的に「グローバルプール」を
//   返すが、毎回 connect() を呼ぶと健全性チェックが走ってオーバーヘッドに
//   なる。ここで ConnectionPool を自分で保持し、再利用する。
//
//   ★ 自己回復のしくみ:
//     - pool.on('error') が発火したら、その時点で破綻しているとみなして
//       明示的に close() し、内部状態を完全クリア。
//     - 次回 getPool() 呼び出しで自動的に新規プールが作成される。
//     - acquireTimeoutMillis: 8000 と組み合わせて、ユーザーの待ち時間を
//       最長でも 8 秒（+ 再接続 10 秒 = 約 18 秒）に抑える。
//
//   旧版では _pool = null するだけで close() しなかったため、
//   破綻したプールがバックグラウンドで居座り続けてリソースを消費していた。
// ============================================================
let _pool = null;
let _poolConnectingPromise = null;

// 破綻したプールを安全に破棄する（エラー検出時のクリーンアップ）
function destroyPool(p, reason) {
    if (!p) return;
    if (_pool === p) _pool = null;
    // close() は失敗しうるが、捨てる前提なので無視する
    p.close().catch(err => {
        console.warn(`🧹 破綻プールの close() 失敗 (${reason}):`, err.message);
    });
}

async function getPool() {
    if (_pool && _pool.connected) return _pool;
    if (_poolConnectingPromise) return _poolConnectingPromise;

    _poolConnectingPromise = (async () => {
        const p = new sql.ConnectionPool(dbConfig);
        p.on('error', err => {
            console.error('🛑 ConnectionPool エラー:', err.message);
            // 旧版では _pool = null するだけだったが、それだと壊れた接続が
            // プール内に残り、後続リクエストが再びそれを掴んでタイムアウト
            // するという「破綻状態の継続」が起きる。明示的に close() して
            // 内部のソケットも解放する。
            destroyPool(p, 'pool error event');
        });
        await p.connect();
        _pool = p;
        _poolConnectingPromise = null;
        return p;
    })();

    try {
        return await _poolConnectingPromise;
    } catch (e) {
        _poolConnectingPromise = null;
        throw e;
    }
}

// 起動時に接続確認
(async () => {
    try {
        await getPool();
        console.log('✅ SQL Server プールを確立しました (192.168.1.8\\SQLEXPRESS)');
    } catch (err) {
        console.error('❌ SQL Server 接続失敗:', err.message);
    }
})();

// プロセス終了時にプールをクローズ
async function shutdownPool() {
    try {
        if (_pool) {
            await _pool.close();
            _pool = null;
            console.log('🧹 ConnectionPool をクローズしました');
        }
    } catch (e) { /* ignore */ }
}
process.on('SIGINT',  async () => { await shutdownPool(); process.exit(0); });
process.on('SIGTERM', async () => { await shutdownPool(); process.exit(0); });

// ============================================================
// 0-a. GET /api/auth/users
//      ログイン画面用（公開）— ユーザー一覧 + PIN 設定済みフラグ
// ============================================================
app.get('/api/auth/users', async (req, res) => {
    try {
        const pool = await getPool();
        const r = await pool.request().query(`
            SELECT
                u.UserID         AS id,
                u.Name           AS name,
                d.DepartmentName AS department,
                CASE WHEN u.PinHash IS NULL THEN 0 ELSE 1 END AS hasPin
            FROM Users AS u
            INNER JOIN Departments AS d ON d.DepartmentID = u.DepartmentID
            WHERE u.IsActive = 1
            ORDER BY d.SortOrder, u.Name
        `);
        res.json(r.recordset.map(row => ({
            id:         row.id,
            name:       row.name,
            department: row.department,
            hasPin:     row.hasPin === 1
        })));
    } catch (err) {
        console.error('/api/auth/users エラー:', err);
        res.status(500).json({ error: 'Fetch error', details: err.message });
    }
});

// ============================================================
// 0-b. POST /api/auth/login
//      body: { userId, pin }
//      - PinHash が NULL の場合は needsSetup:true を返す
//      - 成功時は token + user 情報を返す
// ============================================================
app.post('/api/auth/login', async (req, res) => {
    try {
        const { userId, pin } = req.body || {};
        if (!userId || !isValidPin(pin)) {
            return res.status(400).json({ error: 'ユーザーとPIN（4桁）を入力してください' });
        }
        const pool = await getPool();
        const r = await pool.request()
            .input('UserID', sql.Int, userId)
            .query(`
                SELECT u.UserID, u.Name, u.Role, u.PinHash, u.TokenVersion, u.IsActive,
                       d.DepartmentName
                FROM Users AS u
                INNER JOIN Departments AS d ON d.DepartmentID = u.DepartmentID
                WHERE u.UserID = @UserID
            `);
        const row = r.recordset[0];
        if (!row || !row.IsActive) {
            return res.status(404).json({ error: 'ユーザーが見つかりません' });
        }
        if (row.PinHash === null) {
            return res.status(200).json({ needsSetup: true, userId: row.UserID, name: row.Name });
        }
        if (!verifyPin(pin, row.PinHash)) {
            return res.status(401).json({ error: 'PINが正しくありません' });
        }
        const token = createToken(row.UserID, row.Role, row.TokenVersion);
        res.json({
            token,
            user: {
                id:         row.UserID,
                name:       row.Name,
                role:       row.Role,
                department: row.DepartmentName
            }
        });
    } catch (err) {
        console.error('/api/auth/login エラー:', err);
        res.status(500).json({ error: 'Login error', details: err.message });
    }
});

// ============================================================
// 0-c. POST /api/auth/setup-pin
//      body: { userId, pin }
//      - PinHash が NULL のユーザーのみ設定可能（初回PIN登録）
// ============================================================
app.post('/api/auth/setup-pin', async (req, res) => {
    try {
        const { userId, pin } = req.body || {};
        if (!userId || !isValidPin(pin)) {
            return res.status(400).json({ error: 'PINは4桁の数字で入力してください' });
        }
        const pool = await getPool();
        const r = await pool.request()
            .input('UserID', sql.Int, userId)
            .query(`
                SELECT u.UserID, u.Name, u.Role, u.PinHash, u.IsActive, d.DepartmentName
                FROM Users AS u
                INNER JOIN Departments AS d ON d.DepartmentID = u.DepartmentID
                WHERE u.UserID = @UserID
            `);
        const row = r.recordset[0];
        if (!row || !row.IsActive) {
            return res.status(404).json({ error: 'ユーザーが見つかりません' });
        }
        if (row.PinHash !== null) {
            return res.status(400).json({ error: 'このユーザーは既にPINが設定されています。通常のログインをしてください' });
        }
        const pinHash = hashPin(pin);
        const up = await pool.request()
            .input('UserID',  sql.Int,          row.UserID)
            .input('PinHash', sql.NVarChar(255), pinHash)
            .query(`
                UPDATE Users
                   SET PinHash      = @PinHash,
                       PinSetAt     = GETDATE(),
                       TokenVersion = TokenVersion + 1
                 OUTPUT INSERTED.TokenVersion
                 WHERE UserID = @UserID
            `);
        const newVersion = up.recordset[0].TokenVersion;
        const token = createToken(row.UserID, row.Role, newVersion);
        res.json({
            token,
            user: {
                id:         row.UserID,
                name:       row.Name,
                role:       row.Role,
                department: row.DepartmentName
            }
        });
    } catch (err) {
        console.error('/api/auth/setup-pin エラー:', err);
        res.status(500).json({ error: 'Setup error', details: err.message });
    }
});

// ============================================================
// 0-d. POST /api/auth/change-pin
//      要認証。body: { oldPin, newPin }
// ============================================================
app.post('/api/auth/change-pin', authMiddleware, async (req, res) => {
    try {
        const { oldPin, newPin } = req.body || {};
        if (!isValidPin(oldPin) || !isValidPin(newPin)) {
            return res.status(400).json({ error: 'PINは4桁の数字で入力してください' });
        }
        const pool = await getPool();
        const r = await pool.request()
            .input('UserID', sql.Int, req.user.userId)
            .query('SELECT PinHash FROM Users WHERE UserID = @UserID');
        const row = r.recordset[0];
        if (!row || !row.PinHash) {
            return res.status(400).json({ error: 'PIN が未設定です' });
        }
        if (!verifyPin(oldPin, row.PinHash)) {
            return res.status(401).json({ error: '現在のPINが正しくありません' });
        }
        const newHash = hashPin(newPin);
        const up = await pool.request()
            .input('UserID',  sql.Int,          req.user.userId)
            .input('PinHash', sql.NVarChar(255), newHash)
            .query(`
                UPDATE Users
                   SET PinHash      = @PinHash,
                       PinSetAt     = GETDATE(),
                       TokenVersion = TokenVersion + 1
                 OUTPUT INSERTED.TokenVersion
                 WHERE UserID = @UserID
            `);
        const newVersion = up.recordset[0].TokenVersion;
        const token = createToken(req.user.userId, req.user.role, newVersion);
        res.json({ success: true, token });
    } catch (err) {
        console.error('/api/auth/change-pin エラー:', err);
        res.status(500).json({ error: 'Change PIN error', details: err.message });
    }
});

// ============================================================
// 0-e. POST /api/auth/reset-pin
//      admin のみ。body: { userId }
//      対象ユーザーの PinHash を NULL にし、TokenVersion を進めて強制ログアウト
// ============================================================
app.post('/api/auth/reset-pin', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const { userId } = req.body || {};
        if (!userId) {
            return res.status(400).json({ error: '対象ユーザーが指定されていません' });
        }
        const pool = await getPool();
        const up = await pool.request()
            .input('UserID', sql.Int, userId)
            .query(`
                UPDATE Users
                   SET PinHash      = NULL,
                       PinSetAt     = NULL,
                       TokenVersion = TokenVersion + 1
                 WHERE UserID = @UserID
            `);
        if (up.rowsAffected[0] === 0) {
            return res.status(404).json({ error: 'ユーザーが見つかりません' });
        }
        res.json({ success: true });
    } catch (err) {
        console.error('/api/auth/reset-pin エラー:', err);
        res.status(500).json({ error: 'Reset error', details: err.message });
    }
});

// ============================================================
// 0-f. GET /api/auth/me
//      要認証。現在のログインユーザー情報を返す（トークン再検証用）
// ============================================================
app.get('/api/auth/me', authMiddleware, async (req, res) => {
    try {
        const pool = await getPool();
        const r = await pool.request()
            .input('UserID', sql.Int, req.user.userId)
            .query(`
                SELECT u.UserID, u.Name, u.Role, d.DepartmentName
                FROM Users AS u
                INNER JOIN Departments AS d ON d.DepartmentID = u.DepartmentID
                WHERE u.UserID = @UserID
            `);
        const row = r.recordset[0];
        if (!row) return res.status(404).json({ error: 'ユーザーが見つかりません' });
        res.json({
            id:         row.UserID,
            name:       row.Name,
            role:       row.Role,
            department: row.DepartmentName
        });
    } catch (err) {
        console.error('/api/auth/me エラー:', err);
        res.status(500).json({ error: 'Fetch error', details: err.message });
    }
});

// ============================================================
// 1. GET /api/init-data
//    ユーザー・注番・作業記録・作業内容マスタ・部署・目標時間 一括取得
// ============================================================
app.get('/api/init-data', authMiddleware, async (req, res) => {
    try {
        const pool = await getPool();

        // ★ ログの初期取得範囲
        //   クライアントが dateFrom/dateTo を指定すればそれを使い、
        //   未指定なら「今月1日〜今日」を採用（メモリ節約）。
        //   過去データを見たいときは /api/logs で範囲を広げて再取得する。
        const dateFrom = typeof req.query.dateFrom === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(req.query.dateFrom)
            ? req.query.dateFrom : null;
        const dateTo   = typeof req.query.dateTo === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(req.query.dateTo)
            ? req.query.dateTo   : null;

        // 部署一覧
        const departmentsResult = await pool.request().query(`
            SELECT
                DepartmentID   AS id,
                DepartmentName AS name,
                DepartmentCode AS code,
                SortOrder      AS sortOrder
            FROM Departments
            WHERE IsActive = 1
            ORDER BY SortOrder, DepartmentName
        `);

        // ユーザー一覧（部署名付き）
        const usersResult = await pool.request().query(`
            SELECT
                u.UserID         AS id,
                u.Name           AS name,
                d.DepartmentName AS department,
                u.Email          AS email,
                u.Role           AS role,
                CASE WHEN u.PinHash IS NULL THEN 0 ELSE 1 END AS hasPin
            FROM Users       AS u
            INNER JOIN Departments AS d ON d.DepartmentID = u.DepartmentID
            WHERE u.IsActive = 1
            ORDER BY d.SortOrder, u.Name
        `);

        // 注番一覧（最終使用日が新しい順）
        const projectsResult = await pool.request().query(`
            SELECT
                ProjectNo  AS [no],
                ClientName AS client,
                Subject    AS subject,
                IsClosed   AS isClosed,
                ClosedAt   AS closedAt,
                ClosedReason AS closedReason,
                DATEDIFF_BIG(millisecond, '1970-01-01', ISNULL(LastUsedDate, CreatedAt)) AS lastUsed
            FROM Projects
            ORDER BY LastUsedDate DESC
        `);

        // 作業記録: 指定範囲（未指定なら当月1日〜今日）
        // 注番・入力データが増え続けても初期ロードが重くならないよう、
        // 「最新500件の全期間」ではなく「当月分のみ」を返すのがデフォルト。
        const logsRequest = pool.request();
        if (dateFrom) logsRequest.input('DateFrom', sql.Date, dateFrom);
        if (dateTo)   logsRequest.input('DateTo',   sql.Date, dateTo);
        const logsSql = `
            SELECT
                w.LogID           AS id,
                w.UserID          AS userId,
                w.DepartmentID    AS departmentId,
                d.DepartmentName  AS department,
                w.ProjectNo       AS projectNo,
                w.ContentName     AS content,
                FORMAT(w.WorkDate, 'yyyy-MM-dd') AS date,
                w.WorkHours       AS hours,
                w.WorkLocation    AS location,
                w.IsAfterShipment AS afterShipment,
                w.Details         AS details
            FROM WorkLogs AS w
            INNER JOIN Departments AS d ON d.DepartmentID = w.DepartmentID
            WHERE w.IsDeleted = 0
              AND w.WorkDate >= ${dateFrom ? '@DateFrom' : "DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)"}
              ${dateTo ? 'AND w.WorkDate <= @DateTo' : ''}
            ORDER BY w.WorkDate DESC, w.LogID DESC
        `;
        const logsResult = await logsRequest.query(logsSql);

        // 全作業記録の件数（論理削除を除く）
        const countResult = await pool.request().query(`
            SELECT COUNT(*) AS total FROM WorkLogs WHERE IsDeleted = 0
        `);

        // 取得した範囲内の件数（「もっと読み込む」のページネーション基準）
        const loadedCountRequest = pool.request();
        if (dateFrom) loadedCountRequest.input('DateFrom', sql.Date, dateFrom);
        if (dateTo)   loadedCountRequest.input('DateTo',   sql.Date, dateTo);
        const loadedCountResult = await loadedCountRequest.query(`
            SELECT COUNT(*) AS total
            FROM WorkLogs
            WHERE IsDeleted = 0
              AND WorkDate >= ${dateFrom ? '@DateFrom' : "DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)"}
              ${dateTo ? 'AND WorkDate <= @DateTo' : ''}
        `);

        // 作業内容マスタ（部署ごと）
        const workTypesResult = await pool.request().query(`
            SELECT
                d.DepartmentName AS department,
                wt.TypeName      AS typeName,
                wt.SortOrder     AS sortOrder
            FROM WorkTypes   AS wt
            INNER JOIN Departments AS d ON d.DepartmentID = wt.DepartmentID
            WHERE wt.IsActive = 1
            ORDER BY d.SortOrder, wt.SortOrder
        `);

        // 目標時間マスタ
        const targetsResult = await pool.request().query(`
            SELECT
                TargetID    AS id,
                UserID      AS userId,
                ProjectNo   AS projectNo,
                TargetHours AS targetHours
            FROM ProjectTargets
        `);

        // 作業内容を { 部署名: [内容名, ...] } の形に変換
        const workContents = {};
        for (const row of workTypesResult.recordset) {
            if (!workContents[row.department]) workContents[row.department] = [];
            workContents[row.department].push(row.typeName);
        }

        // 取得範囲（クライアントが知る必要があるため返す）
        const rangeFrom = dateFrom || (() => {
            const d = new Date();
            return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-01`;
        })();
        const rangeTo = dateTo || (() => {
            const d = new Date();
            return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
        })();

        res.json({
            departments:  departmentsResult.recordset.map(d => d.name),
            users:        usersResult.recordset.map(u => ({
                id:         u.id,
                name:       u.name,
                department: u.department,
                email:      u.email,
                role:       u.role,
                hasPin:     u.hasPin === 1
            })),
            projects:     projectsResult.recordset.map(p => ({
                no:           p.no,
                client:       p.client,
                subject:      p.subject,
                isClosed:     !!p.isClosed,
                closedAt:     p.closedAt,
                closedReason: p.closedReason,
                lastUsed:     p.lastUsed
            })),
            logs:         logsResult.recordset,
            totalLogs:    countResult.recordset[0].total,
            rangeTotal:   loadedCountResult.recordset[0].total,
            loadedRange:  { from: rangeFrom, to: rangeTo },
            workContents: workContents,
            targets:      targetsResult.recordset
        });
    } catch (err) {
        console.error('/api/init-data エラー:', err);
        res.status(500).json({ error: 'Database error', details: err.message });
    }
});

// ============================================================
// 2. POST /api/logs
//    作業記録を登録（初回注番の場合は Projects / 初回目標の場合は ProjectTargets にも登録）
// ============================================================
app.post('/api/logs', authMiddleware, async (req, res) => {
    // ---- パフォーマンス計測（一時的な診断ログ） ----
    const t0 = Date.now();
    const marks = {};
    const mark = (label) => { marks[label] = Date.now() - t0; };
    try {
        const data = req.body;
        // なりすまし防止: 自分以外のユーザーIDで記録できない
        if (Number(data.userId) !== Number(req.user.userId)) {
            return res.status(403).json({ error: '他のユーザーとして作業記録を登録することはできません' });
        }
        // 注番の有無判定（空文字/undefined/null はすべて「注番なし」扱い）
        const hasProject = data.projectNo != null && String(data.projectNo).trim() !== '';
        const projectNo  = hasProject ? String(data.projectNo).trim() : null;

        // 注番なし作業は 内容詳細 必須
        if (!hasProject && (!data.details || String(data.details).trim() === '')) {
            return res.status(400).json({ error: '注番なしの作業記録は内容詳細が必須です。' });
        }

        const pool = await getPool();
        mark('getPool');
        const tx   = new sql.Transaction(pool);
        await tx.begin();
        mark('tx.begin');

        try {
            if (hasProject) {
                // 注番の存在確認（クローズ状態も取得）
                const checkProj = await new sql.Request(tx)
                    .input('ProjectNo', sql.NVarChar(50), projectNo)
                    .query('SELECT ProjectNo, IsClosed FROM Projects WHERE ProjectNo = @ProjectNo');
                mark('checkProj');

                if (checkProj.recordset.length === 0) {
                    // 初回注番 → Projects に登録
                    await new sql.Request(tx)
                        .input('ProjectNo',  sql.NVarChar(50),  projectNo)
                        .input('ClientName', sql.NVarChar(100), data.newClient  || '')
                        .input('Subject',    sql.NVarChar(200), data.newSubject || '')
                        .query(`
                            INSERT INTO Projects (ProjectNo, ClientName, Subject, LastUsedDate)
                            VALUES (@ProjectNo, @ClientName, @Subject, GETDATE())
                        `);
                    mark('insertProject');
                } else {
                    // ★ クローズ済み注番への入力は拒否
                    if (checkProj.recordset[0].IsClosed) {
                        await tx.rollback();
                        return res.status(409).json({
                            error: 'この注番はクロースされています。管理者に解除を依頼してください',
                            code:  'PROJECT_CLOSED'
                        });
                    }
                    // 既存注番 → 最終使用日を更新
                    await new sql.Request(tx)
                        .input('ProjectNo', sql.NVarChar(50), projectNo)
                        .query('UPDATE Projects SET LastUsedDate = GETDATE() WHERE ProjectNo = @ProjectNo');
                    mark('updateProject');
                }
            }

            // 目標時間（注番あり かつ 初回のみ渡される）
            let targetRecord = null;
            if (hasProject && data.targetHours && Number(data.targetHours) > 0) {
                const tResult = await new sql.Request(tx)
                    .input('UserID',      sql.Int,         data.userId)
                    .input('ProjectNo',   sql.NVarChar(50),projectNo)
                    .input('TargetHours', sql.Decimal(7,2),Number(data.targetHours))
                    .query(`
                        INSERT INTO ProjectTargets (UserID, ProjectNo, TargetHours)
                        OUTPUT INSERTED.TargetID, INSERTED.UserID, INSERTED.ProjectNo, INSERTED.TargetHours
                        VALUES (@UserID, @ProjectNo, @TargetHours)
                    `);
                targetRecord = tResult.recordset[0];
                mark('insertTarget');
            }

            // 記録時点の所属部署をサーバー側で解決
            const deptResolve = await new sql.Request(tx)
                .input('UserID', sql.Int, data.userId)
                .query('SELECT DepartmentID FROM Users WHERE UserID = @UserID AND IsActive = 1');
            mark('resolveDept');
            if (deptResolve.recordset.length === 0) {
                await tx.rollback();
                return res.status(400).json({ error: 'ユーザーが見つかりません（または無効化されています）' });
            }
            const snapshotDeptId = deptResolve.recordset[0].DepartmentID;

            // 作業記録を挿入
            const result = await new sql.Request(tx)
                .input('ProjectNo',      sql.NVarChar(50),    projectNo)
                .input('WorkDate',       sql.Date,            data.date)
                .input('UserID',         sql.Int,             data.userId)
                .input('DepartmentID',   sql.Int,             snapshotDeptId)
                .input('ContentName',    sql.NVarChar(100),   data.content)
                .input('WorkHours',      sql.Decimal(5, 2),   data.hours)
                .input('WorkLocation',   sql.NVarChar(10),    data.location    || '社内')
                .input('IsAfterShipment',sql.Bit,             data.afterShipment ? 1 : 0)
                .input('Details',        sql.NVarChar(sql.MAX), data.details   || '')
                .query(`
                    INSERT INTO WorkLogs
                        (ProjectNo, WorkDate, UserID, DepartmentID, ContentName, WorkHours, WorkLocation, IsAfterShipment, Details)
                    OUTPUT INSERTED.LogID
                    VALUES
                        (@ProjectNo, @WorkDate, @UserID, @DepartmentID, @ContentName, @WorkHours, @WorkLocation, @IsAfterShipment, @Details)
                `);
            mark('insertWorkLog');

            await tx.commit();
            mark('tx.commit');

            const total = Date.now() - t0;
            console.log(`⏱  POST /api/logs ${total}ms  ${JSON.stringify(marks)}`);

            res.json({
                success: true,
                logId: result.recordset[0].LogID,
                target: targetRecord ? {
                    id: targetRecord.TargetID,
                    userId: targetRecord.UserID,
                    projectNo: targetRecord.ProjectNo,
                    targetHours: Number(targetRecord.TargetHours)
                } : null
            });

        } catch (innerErr) {
            await tx.rollback();
            throw innerErr;
        }
    } catch (err) {
        const total = Date.now() - t0;
        console.error(`⏱  POST /api/logs FAILED ${total}ms  ${JSON.stringify(marks)}`);
        console.error('/api/logs POST エラー:', err);
        res.status(500).json({ error: 'Save error', details: err.message });
    }
});

// ============================================================
// 3. PUT /api/logs/:id
//    作業記録を更新
// ============================================================
app.put('/api/logs/:id', authMiddleware, async (req, res) => {
    try {
        const logId = req.params.id;
        const data  = req.body;
        const pool  = await getPool();

        const hasProject = data.projectNo != null && String(data.projectNo).trim() !== '';
        const projectNo  = hasProject ? String(data.projectNo).trim() : null;
        if (!hasProject && (!data.details || String(data.details).trim() === '')) {
            return res.status(400).json({ error: '注番なしの作業記録は内容詳細が必須です。' });
        }

        // ※ DepartmentID は「記録時点の所属」を保持する設計のため、
        //    編集時は意図的に書き換えない（UserID も通常は変更されない想定）
        await pool.request()
            .input('LogID',          sql.BigInt,          logId)
            .input('ProjectNo',      sql.NVarChar(50),    projectNo)
            .input('WorkDate',       sql.Date,            data.date)
            .input('UserID',         sql.Int,             data.userId)
            .input('ContentName',    sql.NVarChar(100),   data.content)
            .input('WorkHours',      sql.Decimal(5, 2),   data.hours)
            .input('WorkLocation',   sql.NVarChar(10),    data.location    || '社内')
            .input('IsAfterShipment',sql.Bit,             data.afterShipment ? 1 : 0)
            .input('Details',        sql.NVarChar(sql.MAX), data.details   || '')
            .query(`
                UPDATE WorkLogs
                SET ProjectNo       = @ProjectNo,
                    WorkDate        = @WorkDate,
                    UserID          = @UserID,
                    ContentName     = @ContentName,
                    WorkHours       = @WorkHours,
                    WorkLocation    = @WorkLocation,
                    IsAfterShipment = @IsAfterShipment,
                    Details         = @Details,
                    UpdatedAt       = GETDATE()
                WHERE LogID = @LogID
            `);

        res.json({ success: true });
    } catch (err) {
        console.error('/api/logs PUT エラー:', err);
        res.status(500).json({ error: 'Update error', details: err.message });
    }
});

// ============================================================
// 3-b. DELETE /api/logs/:id
//      作業記録を論理削除
// ============================================================
app.delete('/api/logs/:id', authMiddleware, async (req, res) => {
    try {
        const logId = req.params.id;
        const pool  = await getPool();

        await pool.request()
            .input('LogID', sql.BigInt, logId)
            .query('UPDATE WorkLogs SET IsDeleted = 1, UpdatedAt = GETDATE() WHERE LogID = @LogID');

        res.json({ success: true });
    } catch (err) {
        console.error('/api/logs DELETE エラー:', err);
        res.status(500).json({ error: 'Delete error', details: err.message });
    }
});

// ============================================================
// 3-b'. GET /api/me/input-status?days=7
//
//   自分の「直近 N 営業日」の入力状況を返す。
//   ホーム画面の「入力忘れ」パネルが使用する。
//
//   - 土日・祝日・会社休業日は営業日から除外して表示しない
//   - 返り値の days[] は古い順（左 → 右で時間経過）
//   - days[*].logged は「その日に自分の WorkLogs 行が1件以上あるか」
//   - days[*].hours は合計時間（見やすさのため 2 桁まで）
//   - today フィールドはサーバータイムゾーンの「今日」（YYYY-MM-DD）
//
//   認証必須。req.user.userId のデータだけを返すので、他人の分は見えない。
// ============================================================
app.get('/api/me/input-status', authMiddleware, async (req, res) => {
    try {
        const days = Math.min(Math.max(parseInt(req.query.days) || 7, 1), 30);
        const businessDays = lastNBusinessDays(days);
        if (businessDays.length === 0) {
            return res.json({ today: formatDateLocal(new Date()), days: [] });
        }

        const pool = await getPool();
        const request = pool.request().input('UserID', sql.Int, req.user.userId);

        // IN 句用のパラメータを組み立てる（SQL インジェクション対策）
        const paramNames = businessDays.map((_, i) => `@D${i}`);
        businessDays.forEach((d, i) => request.input(`D${i}`, sql.Date, d));

        // CONVERT(CHAR(10), WorkDate, 23) で YYYY-MM-DD 文字列として取得。
        // DATE 型を JS Date 経由で扱うとタイムゾーンずれが起きうるため回避。
        const result = await request.query(`
            SELECT CONVERT(CHAR(10), WorkDate, 23) AS WorkDateStr,
                   SUM(WorkHours) AS Hours,
                   COUNT(*) AS LogCount
              FROM WorkLogs
             WHERE UserID = @UserID
               AND IsDeleted = 0
               AND WorkDate IN (${paramNames.join(',')})
             GROUP BY CONVERT(CHAR(10), WorkDate, 23)
        `);

        const hoursByDate = {};
        const countByDate = {};
        for (const row of result.recordset) {
            hoursByDate[row.WorkDateStr] = Number(row.Hours) || 0;
            countByDate[row.WorkDateStr] = Number(row.LogCount) || 0;
        }

        const payload = businessDays.map(d => {
            // 曜日番号（0=日〜6=土）を返してフロントで表示整形に使う
            const [y, m, day] = d.split('-').map(Number);
            const dow = new Date(y, m - 1, day).getDay();
            const h = hoursByDate[d] || 0;
            return {
                date:   d,
                dow,
                hours:  Math.round(h * 100) / 100,
                count:  countByDate[d] || 0,
                logged: (countByDate[d] || 0) > 0
            };
        });

        res.json({
            today: formatDateLocal(new Date()),
            days:  payload
        });
    } catch (err) {
        console.error('/api/me/input-status エラー:', err);
        res.status(500).json({ error: 'Load error', details: err.message });
    }
});

// ============================================================
// 3-c. GET /api/logs
//      作業記録をページネーション付きで取得（追加読み込み用）
// ============================================================
app.get('/api/logs', authMiddleware, async (req, res) => {
    try {
        const limit  = Math.min(parseInt(req.query.limit)  || 500, 5000);
        const offset = parseInt(req.query.offset) || 0;
        const pool   = await getPool();

        // フィルタ（全てオプション）
        //   dateFrom, dateTo : YYYY-MM-DD の範囲
        //   userId           : 担当者
        //   dept             : 部署名（記録時点のスナップショット）
        //   projectNoQ       : 注番の部分一致（LIKE）
        //   projectMode      : 'has' | 'none'
        const dateFrom    = /^\d{4}-\d{2}-\d{2}$/.test(req.query.dateFrom || '') ? req.query.dateFrom : null;
        const dateTo      = /^\d{4}-\d{2}-\d{2}$/.test(req.query.dateTo   || '') ? req.query.dateTo   : null;
        const userId      = req.query.userId ? parseInt(req.query.userId) : null;
        const dept        = req.query.dept || null;
        const projectNoQ  = (req.query.projectNoQ || '').toString().trim();
        const projectMode = req.query.projectMode || '';

        const wheres = ['w.IsDeleted = 0'];
        const r = pool.request();
        if (dateFrom)   { wheres.push('w.WorkDate >= @DateFrom'); r.input('DateFrom', sql.Date, dateFrom); }
        if (dateTo)     { wheres.push('w.WorkDate <= @DateTo');   r.input('DateTo',   sql.Date, dateTo);   }
        if (userId)     { wheres.push('w.UserID = @UserID');      r.input('UserID',   sql.Int,  userId);   }
        if (dept)       { wheres.push('d.DepartmentName = @Dept'); r.input('Dept',    sql.NVarChar(50), dept); }
        if (projectMode === 'has')  wheres.push('w.ProjectNo IS NOT NULL');
        if (projectMode === 'none') wheres.push('w.ProjectNo IS NULL');
        if (projectNoQ && projectMode !== 'none') {
            wheres.push('w.ProjectNo LIKE @ProjectNoQ');
            r.input('ProjectNoQ', sql.NVarChar(52), `%${projectNoQ}%`);
        }
        const whereClause = wheres.join(' AND ');

        r.input('Limit',  sql.Int, limit);
        r.input('Offset', sql.Int, offset);

        const result = await r.query(`
            SELECT
                w.LogID           AS id,
                w.UserID          AS userId,
                w.DepartmentID    AS departmentId,
                d.DepartmentName  AS department,
                w.ProjectNo       AS projectNo,
                w.ContentName     AS content,
                FORMAT(w.WorkDate, 'yyyy-MM-dd') AS date,
                w.WorkHours       AS hours,
                w.WorkLocation    AS location,
                w.IsAfterShipment AS afterShipment,
                w.Details         AS details
            FROM WorkLogs AS w
            INNER JOIN Departments AS d ON d.DepartmentID = w.DepartmentID
            WHERE ${whereClause}
            ORDER BY w.WorkDate DESC, w.LogID DESC
            OFFSET @Offset ROWS FETCH NEXT @Limit ROWS ONLY
        `);

        // フィルタ適用時の該当件数
        const c = pool.request();
        if (dateFrom)   c.input('DateFrom', sql.Date, dateFrom);
        if (dateTo)     c.input('DateTo',   sql.Date, dateTo);
        if (userId)     c.input('UserID',   sql.Int,  userId);
        if (dept)       c.input('Dept',     sql.NVarChar(50), dept);
        if (projectNoQ && projectMode !== 'none') c.input('ProjectNoQ', sql.NVarChar(52), `%${projectNoQ}%`);
        const countResult = await c.query(`
            SELECT COUNT(*) AS total
            FROM WorkLogs AS w
            INNER JOIN Departments AS d ON d.DepartmentID = w.DepartmentID
            WHERE ${whereClause}
        `);

        res.json({
            logs:  result.recordset,
            total: countResult.recordset[0].total
        });
    } catch (err) {
        console.error('/api/logs GET エラー:', err);
        res.status(500).json({ error: 'Fetch error', details: err.message });
    }
});

// ============================================================
// 3-d. PUT /api/targets
//      目標時間を更新（既存目標の変更は leader/admin のみ、reason 必須）
//      body: { userId, projectNo, targetHours, reason }
// ============================================================
app.put('/api/targets', authMiddleware, async (req, res) => {
    try {
        const { userId, projectNo, targetHours, reason } = req.body;
        if (!userId || !projectNo || !(Number(targetHours) > 0)) {
            return res.status(400).json({ error: '必須パラメータが不足しています' });
        }
        const pool = await getPool();
        const tx   = new sql.Transaction(pool);
        await tx.begin();

        try {
            const existing = await new sql.Request(tx)
                .input('UserID',    sql.Int,          userId)
                .input('ProjectNo', sql.NVarChar(50), projectNo)
                .query('SELECT TargetID, TargetHours FROM ProjectTargets WHERE UserID = @UserID AND ProjectNo = @ProjectNo');

            let record;
            let oldValue = null;
            const isUpdate = existing.recordset.length > 0;

            if (isUpdate) {
                // 既存目標の変更は leader / admin のみ、かつ reason 必須
                if (req.user.role !== 'leader' && req.user.role !== 'admin') {
                    await tx.rollback();
                    return res.status(403).json({
                        error: '目標時間の変更権限がありません（リーダーまたは管理者のみ）',
                        code: 'FORBIDDEN'
                    });
                }
                if (!reason || String(reason).trim() === '') {
                    await tx.rollback();
                    return res.status(400).json({ error: '変更理由は必須です' });
                }
                oldValue = Number(existing.recordset[0].TargetHours);

                // ProjectTargets には AFTER UPDATE トリガー (TR_ProjectTargets_UpdatedAt) が
                // 設定されているため、UPDATE 文では OUTPUT 句（INTO 無し）が使えない。
                // UPDATE と SELECT を分離して対応する。
                await new sql.Request(tx)
                    .input('TargetID',    sql.Int,          existing.recordset[0].TargetID)
                    .input('TargetHours', sql.Decimal(7,2), Number(targetHours))
                    .query(`
                        UPDATE ProjectTargets
                           SET TargetHours = @TargetHours
                         WHERE TargetID = @TargetID
                    `);
                const fetched = await new sql.Request(tx)
                    .input('TargetID', sql.Int, existing.recordset[0].TargetID)
                    .query(`
                        SELECT TargetID, UserID, ProjectNo, TargetHours
                          FROM ProjectTargets
                         WHERE TargetID = @TargetID
                    `);
                record = fetched.recordset[0];
            } else {
                // 初回登録は全ロール可
                const result = await new sql.Request(tx)
                    .input('UserID',      sql.Int,          userId)
                    .input('ProjectNo',   sql.NVarChar(50), projectNo)
                    .input('TargetHours', sql.Decimal(7,2), Number(targetHours))
                    .query(`
                        INSERT INTO ProjectTargets (UserID, ProjectNo, TargetHours)
                        OUTPUT INSERTED.TargetID, INSERTED.UserID, INSERTED.ProjectNo, INSERTED.TargetHours
                        VALUES (@UserID, @ProjectNo, @TargetHours)
                    `);
                record = result.recordset[0];
            }

            // 監査ログ
            const logReason = isUpdate
                ? String(reason).trim()
                : '初回登録';
            await new sql.Request(tx)
                .input('TargetID',  sql.Int,           record.TargetID)
                .input('UserID',    sql.Int,           record.UserID)
                .input('ProjectNo', sql.NVarChar(50),  record.ProjectNo)
                .input('OldValue',  sql.Decimal(7,2),  oldValue)
                .input('NewValue',  sql.Decimal(7,2),  Number(record.TargetHours))
                .input('Reason',    sql.NVarChar(500), logReason)
                .input('ChangedBy', sql.Int,           req.user.userId)
                .query(`
                    INSERT INTO TargetChangeLog
                        (TargetID, UserID, ProjectNo, OldValue, NewValue, Reason, ChangedBy)
                    VALUES
                        (@TargetID, @UserID, @ProjectNo, @OldValue, @NewValue, @Reason, @ChangedBy)
                `);

            await tx.commit();
            res.json({
                success: true,
                target: {
                    id:          record.TargetID,
                    userId:      record.UserID,
                    projectNo:   record.ProjectNo,
                    targetHours: Number(record.TargetHours)
                }
            });
        } catch (innerErr) {
            await tx.rollback();
            throw innerErr;
        }
    } catch (err) {
        console.error('/api/targets PUT エラー:', err);
        res.status(500).json({ error: 'Update error', details: err.message });
    }
});

// ============================================================
// 3-d-2. GET /api/targets/:userId/:projectNo/history
//        目標時間の変更履歴（leader/admin のみ閲覧可）
// ============================================================
app.get('/api/targets/:userId/:projectNo/history', authMiddleware, requireRole('leader', 'admin'), async (req, res) => {
    try {
        const userId    = parseInt(req.params.userId);
        // Express は req.params を既にデコード済みで返すため、二重デコードしない
        const projectNo = req.params.projectNo;
        const pool = await getPool();
        const r = await pool.request()
            .input('UserID',    sql.Int,          userId)
            .input('ProjectNo', sql.NVarChar(50), projectNo)
            .query(`
                SELECT
                    tcl.LogID        AS id,
                    tcl.OldValue     AS oldValue,
                    tcl.NewValue     AS newValue,
                    tcl.Reason       AS reason,
                    FORMAT(tcl.ChangedAt, 'yyyy-MM-dd HH:mm:ss') AS changedAt,
                    u.Name           AS changedByName,
                    tcl.ChangedBy    AS changedBy
                FROM TargetChangeLog AS tcl
                INNER JOIN Users AS u ON u.UserID = tcl.ChangedBy
                WHERE tcl.UserID = @UserID AND tcl.ProjectNo = @ProjectNo
                ORDER BY tcl.ChangedAt DESC, tcl.LogID DESC
            `);
        res.json(r.recordset.map(row => ({
            id:            row.id,
            oldValue:      row.oldValue === null ? null : Number(row.oldValue),
            newValue:      Number(row.newValue),
            reason:        row.reason,
            changedAt:     row.changedAt,
            changedBy:     row.changedBy,
            changedByName: row.changedByName
        })));
    } catch (err) {
        console.error('/api/targets history エラー:', err);
        res.status(500).json({ error: 'Fetch error', details: err.message });
    }
});

// ============================================================
// 3-e. PUT /api/projects/:no
//      注番の客先名・件名・クローズ状態を更新
// ============================================================
app.put('/api/projects/:no', authMiddleware, async (req, res) => {
    try {
        const projectNo = decodeURIComponent(req.params.no);
        const { client, subject, isClosed } = req.body;
        const pool = await getPool();

        // クローズ状態の差分検知のため、既存値を取得
        const cur = await pool.request()
            .input('ProjectNo', sql.NVarChar(50), projectNo)
            .query('SELECT IsClosed FROM Projects WHERE ProjectNo = @ProjectNo');
        if (cur.recordset.length === 0) {
            return res.status(404).json({ error: '注番が見つかりません' });
        }
        const wasClosed = !!cur.recordset[0].IsClosed;
        const willClose = !!isClosed;

        // ClosedAt / ClosedReason の更新ロジック:
        //   - false → true の時点で現在時刻と 'manual' を記録
        //   - true → false（再開）の時点でクリア
        //   - 既に閉じていて内容編集だけの場合は触らない
        let setClosedAt     = null;
        let setClosedReason = null;
        let touchClose      = false;
        if (willClose && !wasClosed) { touchClose = true; setClosedAt = new Date(); setClosedReason = 'manual'; }
        if (!willClose && wasClosed) { touchClose = true; setClosedAt = null;       setClosedReason = null;     }

        const r = pool.request()
            .input('ProjectNo',  sql.NVarChar(50),  projectNo)
            .input('ClientName', sql.NVarChar(100), client  || '')
            .input('Subject',    sql.NVarChar(200), subject || '')
            .input('IsClosed',   sql.Bit,           willClose ? 1 : 0);
        if (touchClose) {
            r.input('ClosedAt',     sql.DateTime2,   setClosedAt);
            r.input('ClosedReason', sql.NVarChar(50), setClosedReason);
        }
        await r.query(`
            UPDATE Projects
               SET ClientName = @ClientName,
                   Subject    = @Subject,
                   IsClosed   = @IsClosed
                   ${touchClose ? ', ClosedAt = @ClosedAt, ClosedReason = @ClosedReason' : ''}
             WHERE ProjectNo = @ProjectNo
        `);
        res.json({ success: true });
    } catch (err) {
        console.error('/api/projects PUT エラー:', err);
        res.status(500).json({ error: 'Update error', details: err.message });
    }
});

// ============================================================
// 4. POST /api/users
//    ユーザーを新規登録
// ============================================================
app.post('/api/users', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const data = req.body;
        const pool = await getPool();

        // 部署IDを名前から取得
        const deptResult = await pool.request()
            .input('DepartmentName', sql.NVarChar(50), data.department)
            .query('SELECT DepartmentID FROM Departments WHERE DepartmentName = @DepartmentName');

        if (deptResult.recordset.length === 0) {
            return res.status(400).json({ error: '指定された部署が見つかりません: ' + data.department });
        }
        const deptId = deptResult.recordset[0].DepartmentID;

        const role = ['member', 'leader', 'admin'].includes(data.role) ? data.role : 'member';

        const result = await pool.request()
            .input('DepartmentID', sql.Int,          deptId)
            .input('Name',         sql.NVarChar(50), data.name)
            .input('Email',        sql.NVarChar(100),data.email || '')
            .input('Role',         sql.NVarChar(20), role)
            .query(`
                INSERT INTO Users (DepartmentID, Name, Email, Role, IsActive)
                OUTPUT INSERTED.UserID
                VALUES (@DepartmentID, @Name, @Email, @Role, 1)
            `);

        res.json({ success: true, userId: result.recordset[0].UserID });
    } catch (err) {
        console.error('/api/users POST エラー:', err);
        res.status(500).json({ error: 'Save error', details: err.message });
    }
});

// ============================================================
// 5. PUT /api/users/:id
//    ユーザー情報を更新
// ============================================================
app.put('/api/users/:id', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const userId = req.params.id;
        const data   = req.body;
        const pool   = await getPool();

        // 部署IDを名前から取得
        const deptResult = await pool.request()
            .input('DepartmentName', sql.NVarChar(50), data.department)
            .query('SELECT DepartmentID FROM Departments WHERE DepartmentName = @DepartmentName');

        if (deptResult.recordset.length === 0) {
            return res.status(400).json({ error: '指定された部署が見つかりません: ' + data.department });
        }
        const deptId = deptResult.recordset[0].DepartmentID;

        const role = ['member', 'leader', 'admin'].includes(data.role) ? data.role : 'member';

        await pool.request()
            .input('UserID',       sql.Int,          userId)
            .input('DepartmentID', sql.Int,          deptId)
            .input('Name',         sql.NVarChar(50), data.name)
            .input('Email',        sql.NVarChar(100),data.email || '')
            .input('Role',         sql.NVarChar(20), role)
            .query(`
                UPDATE Users
                SET DepartmentID = @DepartmentID,
                    Name         = @Name,
                    Email        = @Email,
                    Role         = @Role
                WHERE UserID = @UserID
            `);

        res.json({ success: true });
    } catch (err) {
        console.error('/api/users PUT エラー:', err);
        res.status(500).json({ error: 'Update error', details: err.message });
    }
});

// ============================================================
// 6. DELETE /api/users/:id
//    ユーザーを論理削除（IsActive = 0）
// ============================================================
app.delete('/api/users/:id', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const userId = req.params.id;
        const pool   = await getPool();

        await pool.request()
            .input('UserID', sql.Int, userId)
            .query('UPDATE Users SET IsActive = 0 WHERE UserID = @UserID');

        res.json({ success: true });
    } catch (err) {
        console.error('/api/users DELETE エラー:', err);
        res.status(500).json({ error: 'Delete error', details: err.message });
    }
});

// ============================================================
// 7. PUT /api/work-contents/:dept
//    部署の作業内容マスタを一括更新
// ============================================================
app.put('/api/work-contents/:dept', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const dept  = decodeURIComponent(req.params.dept);
        const items = req.body.items; // 文字列の配列
        const pool  = await getPool();
        const tx    = new sql.Transaction(pool);
        await tx.begin();

        try {
            // 部署IDを取得
            const deptResult = await new sql.Request(tx)
                .input('DepartmentName', sql.NVarChar(50), dept)
                .query('SELECT DepartmentID FROM Departments WHERE DepartmentName = @DepartmentName');

            if (deptResult.recordset.length === 0) {
                await tx.rollback();
                return res.status(400).json({ error: '指定された部署が見つかりません: ' + dept });
            }
            const deptId = deptResult.recordset[0].DepartmentID;

            // 既存の作業内容を一旦すべて無効化
            await new sql.Request(tx)
                .input('DepartmentID', sql.Int, deptId)
                .query('UPDATE WorkTypes SET IsActive = 0 WHERE DepartmentID = @DepartmentID');

            // 送られてきた items を順番通りに登録（既存レコードがあれば有効化、なければ挿入）
            for (let i = 0; i < items.length; i++) {
                const typeName = items[i];

                const existing = await new sql.Request(tx)
                    .input('DepartmentID', sql.Int,          deptId)
                    .input('TypeName',     sql.NVarChar(100), typeName)
                    .query(`
                        SELECT WorkTypeID FROM WorkTypes
                        WHERE DepartmentID = @DepartmentID AND TypeName = @TypeName
                    `);

                if (existing.recordset.length > 0) {
                    // 既存レコードを有効化・順序更新
                    await new sql.Request(tx)
                        .input('WorkTypeID', sql.Int, existing.recordset[0].WorkTypeID)
                        .input('SortOrder',  sql.Int, i + 1)
                        .query('UPDATE WorkTypes SET IsActive = 1, SortOrder = @SortOrder WHERE WorkTypeID = @WorkTypeID');
                } else {
                    // 新規挿入
                    await new sql.Request(tx)
                        .input('DepartmentID', sql.Int,          deptId)
                        .input('TypeName',     sql.NVarChar(100), typeName)
                        .input('SortOrder',    sql.Int,           i + 1)
                        .query(`
                            INSERT INTO WorkTypes (DepartmentID, TypeName, SortOrder, IsActive)
                            VALUES (@DepartmentID, @TypeName, @SortOrder, 1)
                        `);
                }
            }

            await tx.commit();
            res.json({ success: true });

        } catch (innerErr) {
            await tx.rollback();
            throw innerErr;
        }
    } catch (err) {
        console.error('/api/work-contents PUT エラー:', err);
        res.status(500).json({ error: 'Update error', details: err.message });
    }
});

// ============================================================
// 8. POST /api/departments
//    部署（グループ）の新規登録
// ============================================================
app.post('/api/departments', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const data = req.body; // { name, code? }
        const pool = await getPool();

        // 既存確認（論理削除済みは復活させる）
        const existing = await pool.request()
            .input('Name', sql.NVarChar(50), data.name)
            .query('SELECT DepartmentID, IsActive FROM Departments WHERE DepartmentName = @Name');

        if (existing.recordset.length > 0) {
            const row = existing.recordset[0];
            if (row.IsActive) {
                return res.status(400).json({ error: '同名の部署が既に存在します' });
            }
            await pool.request()
                .input('DepartmentID', sql.Int, row.DepartmentID)
                .query('UPDATE Departments SET IsActive = 1 WHERE DepartmentID = @DepartmentID');
            return res.json({ success: true, departmentId: row.DepartmentID });
        }

        // 末尾に追加
        const maxOrder = await pool.request()
            .query('SELECT ISNULL(MAX(SortOrder),0) AS mx FROM Departments');
        const nextOrder = maxOrder.recordset[0].mx + 1;

        const result = await pool.request()
            .input('Name',      sql.NVarChar(50), data.name)
            .input('Code',      sql.NVarChar(20), data.code || null)
            .input('SortOrder', sql.Int,          nextOrder)
            .query(`
                INSERT INTO Departments (DepartmentName, DepartmentCode, SortOrder, IsActive)
                OUTPUT INSERTED.DepartmentID
                VALUES (@Name, @Code, @SortOrder, 1)
            `);
        res.json({ success: true, departmentId: result.recordset[0].DepartmentID });
    } catch (err) {
        console.error('/api/departments POST エラー:', err);
        res.status(500).json({ error: 'Save error', details: err.message });
    }
});

// ============================================================
// 9. PUT /api/departments/:oldName
//    部署名を変更
// ============================================================
app.put('/api/departments/:oldName', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const oldName = decodeURIComponent(req.params.oldName);
        const { newName } = req.body;
        const pool = await getPool();

        const result = await pool.request()
            .input('OldName', sql.NVarChar(50), oldName)
            .input('NewName', sql.NVarChar(50), newName)
            .query(`
                UPDATE Departments
                   SET DepartmentName = @NewName
                 WHERE DepartmentName = @OldName
            `);
        res.json({ success: true });
    } catch (err) {
        console.error('/api/departments PUT エラー:', err);
        res.status(500).json({ error: 'Update error', details: err.message });
    }
});

// ============================================================
// 10. DELETE /api/departments/:name
//     部署を論理削除（ユーザーが所属している場合は拒否）
// ============================================================
app.delete('/api/departments/:name', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const name = decodeURIComponent(req.params.name);
        const pool = await getPool();

        // 所属ユーザーがいないか確認
        const check = await pool.request()
            .input('Name', sql.NVarChar(50), name)
            .query(`
                SELECT COUNT(*) AS cnt
                FROM Users u
                INNER JOIN Departments d ON d.DepartmentID = u.DepartmentID
                WHERE d.DepartmentName = @Name AND u.IsActive = 1
            `);
        if (check.recordset[0].cnt > 0) {
            return res.status(400).json({ error: 'このグループには所属ユーザーがいるため削除できません' });
        }

        await pool.request()
            .input('Name', sql.NVarChar(50), name)
            .query('UPDATE Departments SET IsActive = 0 WHERE DepartmentName = @Name');
        res.json({ success: true });
    } catch (err) {
        console.error('/api/departments DELETE エラー:', err);
        res.status(500).json({ error: 'Delete error', details: err.message });
    }
});

// ============================================================
// 11. 自動クローズ / メンテナンス機能
// ============================================================

// 2ヶ月入力のない進行中の注番を自動ロック
//   - 最終入力日（MAX(WorkDate)）を見る。入力が一度もない注番は
//     CreatedAt を最終入力日とみなす
//   - 既にクローズ済みの注番は対象外
async function autoCloseInactiveProjects(pool) {
    const result = await pool.request()
        .input('Months', sql.Int, AUTO_CLOSE_INACTIVE_MONTHS)
        .query(`
            ;WITH LastTouch AS (
                SELECT
                    p.ProjectNo,
                    COALESCE(
                        (SELECT MAX(w.WorkDate) FROM WorkLogs w
                           WHERE w.ProjectNo = p.ProjectNo AND w.IsDeleted = 0),
                        CAST(p.CreatedAt AS DATE)
                    ) AS LastDate
                FROM Projects p
                WHERE p.IsClosed = 0
            )
            UPDATE p
               SET IsClosed     = 1,
                   ClosedAt     = GETDATE(),
                   ClosedReason = N'auto_inactive'
            OUTPUT INSERTED.ProjectNo
              FROM Projects p
              INNER JOIN LastTouch lt ON lt.ProjectNo = p.ProjectNo
             WHERE p.IsClosed = 0
               AND lt.LastDate < DATEADD(MONTH, -@Months, CAST(GETDATE() AS DATE));
        `);
    return result.recordset.map(r => r.ProjectNo);
}

// 起動時に 1 回 + 24 時間ごとに定期実行
async function runAutoCloseScheduled() {
    try {
        const pool = await getPool();
        const closed = await autoCloseInactiveProjects(pool);
        if (closed.length > 0) {
            console.log(`🔒 自動クローズ: ${closed.length} 件の注番をクローズしました`);
            console.log('   ' + closed.slice(0, 10).join(', ') + (closed.length > 10 ? ' ...' : ''));
        } else {
            console.log(`🔒 自動クローズ: 対象なし（${AUTO_CLOSE_INACTIVE_MONTHS}ヶ月以上未使用の進行中注番なし）`);
        }
    } catch (err) {
        console.error('🔒 自動クローズ実行エラー:', err.message);
    }
}

// 手動トリガ（admin のみ）
app.post('/api/admin/auto-close-projects', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const pool = await getPool();
        const closed = await autoCloseInactiveProjects(pool);
        res.json({ success: true, closedCount: closed.length, closed });
    } catch (err) {
        console.error('/api/admin/auto-close-projects エラー:', err);
        res.status(500).json({ error: 'Auto close error', details: err.message });
    }
});

// ============================================================
// 12. GET /api/stats/yoy  前年同期比較
//     dateFrom, dateTo を受け取り、当該期間と前年同期の集計を返す
//     戻り値: { current: { totalHours, byDept, byProject }, previous: {...} }
// ============================================================
app.get('/api/stats/yoy', authMiddleware, async (req, res) => {
    try {
        const dateFrom = req.query.dateFrom;
        const dateTo   = req.query.dateTo;
        if (!/^\d{4}-\d{2}-\d{2}$/.test(dateFrom || '') || !/^\d{4}-\d{2}-\d{2}$/.test(dateTo || '')) {
            return res.status(400).json({ error: 'dateFrom / dateTo (YYYY-MM-DD) が必要です' });
        }
        const shiftYear = (iso, years) => {
            const [y, m, d] = iso.split('-').map(Number);
            return `${y - years}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
        };
        const prevFrom = shiftYear(dateFrom, 1);
        const prevTo   = shiftYear(dateTo,   1);

        const pool = await getPool();
        const query = `
            SELECT
                CAST(SUM(w.WorkHours) AS DECIMAL(18,2))       AS totalHours,
                COUNT(*)                                      AS logCount,
                COUNT(DISTINCT w.ProjectNo)                   AS projectCount,
                CAST(SUM(CASE WHEN w.IsAfterShipment=1 THEN w.WorkHours ELSE 0 END) AS DECIMAL(18,2)) AS afterHours,
                CAST(SUM(CASE WHEN w.WorkLocation=N'社外'  THEN w.WorkHours ELSE 0 END) AS DECIMAL(18,2)) AS outsideHours
            FROM WorkLogs w
            WHERE w.IsDeleted = 0
              AND w.WorkDate BETWEEN @From AND @To
        `;
        const byDeptQ = `
            SELECT d.DepartmentName AS dept, CAST(SUM(w.WorkHours) AS DECIMAL(18,2)) AS hours
            FROM WorkLogs w
            INNER JOIN Departments d ON d.DepartmentID = w.DepartmentID
            WHERE w.IsDeleted = 0 AND w.WorkDate BETWEEN @From AND @To
            GROUP BY d.DepartmentName
            ORDER BY d.DepartmentName
        `;

        const load = async (from, to) => {
            const t = await pool.request()
                .input('From', sql.Date, from)
                .input('To',   sql.Date, to)
                .query(query);
            const d = await pool.request()
                .input('From', sql.Date, from)
                .input('To',   sql.Date, to)
                .query(byDeptQ);
            const row = t.recordset[0] || {};
            return {
                range:        { from, to },
                totalHours:   Number(row.totalHours   || 0),
                logCount:     Number(row.logCount     || 0),
                projectCount: Number(row.projectCount || 0),
                afterHours:   Number(row.afterHours   || 0),
                outsideHours: Number(row.outsideHours || 0),
                byDept:       d.recordset.map(r => ({ dept: r.dept, hours: Number(r.hours || 0) }))
            };
        };

        const current  = await load(dateFrom, dateTo);
        const previous = await load(prevFrom, prevTo);
        res.json({ current, previous });
    } catch (err) {
        console.error('/api/stats/yoy エラー:', err);
        res.status(500).json({ error: 'Fetch error', details: err.message });
    }
});

// ============================================================
// 13. POST /api/admin/maintenance
//     定期保守を一括実行（admin のみ）
//     - 注番の自動クローズ
//     - 古い TargetChangeLog の整理（5年保持）
//     - ※ 認証トークンはステートレス HMAC のため DB クリーンアップ不要。
//       ただし長期間ログインしていないユーザーのセッション強制失効が
//       必要な場合はここで TokenVersion をインクリメントする拡張点。
// ============================================================
app.post('/api/admin/maintenance', authMiddleware, requireRole('admin'), async (req, res) => {
    try {
        const pool = await getPool();

        // 1) 自動クローズ
        const closed = await autoCloseInactiveProjects(pool);

        // 2) 古い監査ログの整理（5年より古いもの）
        const purge = await pool.request().query(`
            DELETE FROM TargetChangeLog
             WHERE ChangedAt < DATEADD(YEAR, -5, GETDATE())
        `);

        // 3) 統計情報の更新（クエリ最適化のヒント）
        await pool.request().query(`
            UPDATE STATISTICS WorkLogs;
            UPDATE STATISTICS Projects;
            UPDATE STATISTICS ProjectTargets;
        `);

        res.json({
            success: true,
            autoClosedCount:          closed.length,
            autoClosedProjects:       closed,
            purgedTargetChangeLogs:   purge.rowsAffected[0] || 0
        });
    } catch (err) {
        console.error('/api/admin/maintenance エラー:', err);
        res.status(500).json({ error: 'Maintenance error', details: err.message });
    }
});

// ============================================================
// サーバー起動
// ============================================================
app.listen(PORT, '0.0.0.0', async () => {
    console.log('');
    console.log('🚀 APIサーバーが起動しました');
    console.log(`🔗 ローカル: http://localhost:${PORT}`);
    console.log(`🔗 外部アクセス: http://192.168.1.8:${PORT}`);
    console.log('');

    // 起動時に 1 度 + 24 時間ごとに自動クローズを実行
    setTimeout(runAutoCloseScheduled, 5000);
    setInterval(runAutoCloseScheduled, AUTO_CLOSE_INTERVAL_MS);
});

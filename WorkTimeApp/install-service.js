// ============================================================
// Windowsサービス インストールスクリプト
// 
// 【実行方法】
//   node install-service.js
// ============================================================

const Service = require('node-windows').Service;
const path    = require('path');

// ★ここだけ自分の環境に合わせて変更してください★
const SERVER_JS_PATH = path.join(__dirname, 'server.js');

const svc = new Service({
    name:        '作業時間管理システム',          // サービス名（日本語OK）
    description: '作業時間 一元管理システム APIサーバー',
    script:      SERVER_JS_PATH,
    nodeOptions: [],

    // サービスが異常終了した場合の自動再起動設定
    wait:   2,   // 再起動まで2秒待つ
    grow:   0.5, // 再起動ごとに待機時間を1.5倍に増やす
    maxRestarts: 5  // 最大5回まで自動再起動
});

// インストール完了イベント
svc.on('install', () => {
    console.log('✅ サービスのインストールが完了しました！');
    console.log('🚀 サービスを起動します...');
    svc.start();
});

// 起動完了イベント
svc.on('start', () => {
    console.log('✅ サービスが起動しました！');
    console.log('');
    console.log('確認方法:');
    console.log('  - Windowsサービス画面で「作業時間管理システム」を探してください');
    console.log(`  - ブラウザで http://localhost:3000/api/init-data を開いて確認`);
});

svc.on('error', (err) => {
    console.error('❌ エラーが発生しました:', err);
});

// インストール実行
console.log('サービスをインストール中...');
console.log('対象ファイル:', SERVER_JS_PATH);
svc.install();

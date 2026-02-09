
function debugLineProps() {
  const sp = PropertiesService.getScriptProperties();
  const all = sp.getProperties();
  Logger.log('All Script Properties:\n' + JSON.stringify(all, null, 2));
  Logger.log('Has CHANNEL_TOKEN? ' + Boolean(sp.getProperty('LINE_CHANNEL_ACCESS_TOKEN')));
  Logger.log('Has ADMIN_USER_ID? ' + Boolean(sp.getProperty('LINE_ADMIN_USER_ID')));
  Logger.log('Has NOTIFY_TOKEN?  ' + Boolean(sp.getProperty('LINE_NOTIFY_ACCESS_TOKEN')));
  Logger.log('Script ID: ' + ScriptApp.getScriptId());
}

function setAdminUserIdAndVerify() {
  const KEY = 'LINE_ADMIN_USER_ID';
  const VALUE = 'U3416fc5085d25c254fe6164ff3e20cb6'; // ← 実際の U から始まるID

  const sp = PropertiesService.getScriptProperties();

  // 上書き保存
  sp.setProperty(KEY, String(VALUE).trim());

  // すぐ読み戻して確認
  const after = sp.getProperty(KEY);
  Logger.log('Saved ADMIN_USER_ID? ' + (after ? 'YES: ' + after : 'NO'));

  // 登録済みキー一覧も確認
  const all = sp.getProperties();
  Logger.log('All keys: ' + Object.keys(all).join(', '));
}

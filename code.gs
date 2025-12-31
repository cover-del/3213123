const CONFIG = {
  ROOM_KEY: 'WEREWOLF_ROOMS',
  AVATAR_FOLDER: 'Werewolf_Avatars',
  POLL_INTERVAL_MS: 1500,
  DEFAULT_PLAYERS: 6,
  ROLE_DISTRIBUTION: ['werewolf','werewolf','seer','doctor','villager','villager'],
  // 新增 Google Sheet 設定
  SHEET_ID: '1VHkP7WhIEaLQ-S87AcNq_R8hA2Hz1XTjyC6md1GM3LI', // 使用者需填入試算表 ID，若留空則自動建立
  SHEET_NAME: 'RoomsData'
};

const ADMIN_PASSWORD = '1234';

// ===== 基礎工具 (修改為 Google Sheets 儲存) =====

/**
 * 取得或建立用於儲存資料的試算表
 */
function _getOrCreateSheet() {
  let ss;
  const scriptProps = PropertiesService.getScriptProperties();
  let ssId = CONFIG.SHEET_ID || scriptProps.getProperty('DB_SHEET_ID');
  
  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ss = null;
    }
  }
  
  if (!ss) {
    ss = SpreadsheetApp.create('Werewolf_Game_DB');
    ssId = ss.getId();
    scriptProps.setProperty('DB_SHEET_ID', ssId);
  }
  
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // 初始化標題列
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _loadRooms() {
  // 優先嘗試從 PropertiesService 讀取以保持效能，若無則從 Sheet 讀取
  const p = PropertiesService.getScriptProperties().getProperty(CONFIG.ROOM_KEY);
  if (p) return JSON.parse(p);
  
  // 從 Sheet 讀取
  const sheet = _getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === CONFIG.ROOM_KEY) {
      const val = data[i][1];
      // 同步回 PropertiesService
      PropertiesService.getScriptProperties().setProperty(CONFIG.ROOM_KEY, val);
      return val ? JSON.parse(val) : {};
    }
  }
  return {};
}

function _saveRooms(obj) {
  const jsonStr = JSON.stringify(obj);
  // 同步到 PropertiesService (快速讀取)
  PropertiesService.getScriptProperties().setProperty(CONFIG.ROOM_KEY, jsonStr);
  
  // 同步到 Google Sheet (持久化儲存)
  const sheet = _getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === CONFIG.ROOM_KEY) {
      sheet.getRange(i + 1, 2).setValue(jsonStr);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([CONFIG.ROOM_KEY, jsonStr]);
  }
}

function _now(){ return (new Date()).toISOString(); }
function _uid(prefix){ return (prefix||'id_') + Math.random().toString(36).slice(2,9); }
function _shuffle(arr){ for(let i=arr.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [arr[i],arr[j]]=[arr[j],arr[i]]; } return arr; }
function _getAvatarFolder(){
  const name = CONFIG.AVATAR_FOLDER;
  const it = DriveApp.getFoldersByName(name);
  if(it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

// ===== Web 前端 =====
/**
 * Web App 入口
 * 支援三種頁面：
 * - login: 登入頁面（預設）
 * - index: 遊戲主頁
 * - admin: 管理頁面（需 ?admin=1）
 */
function doGet(e) {
  const params = e.parameter || {};
  const isAdmin = params.admin === '1';
  const page = params.page || 'login'; // 預設 login

  // 確定要載入的 HTML
  let htmlFile = 'login'; // 預設
  if (isAdmin) {
    htmlFile = 'admin';
  } else if (page === 'index') {
    htmlFile = 'index';
  }

  // 動態設定標題
  const titleMap = {
    'login': '狼人殺登入',
    'index': '狼人殺遊戲',
    'admin': '狼人殺 Admin 管理'
  };
  const title = titleMap[htmlFile] || '狼人殺';

  // 建立模板
  const template = HtmlService.createTemplateFromFile(htmlFile);

  // 可加入共用變數，例如 API_URL、版本號等
  template.API_URL = ScriptApp.getService().getUrl(); // Web App URL
  template.VERSION = '1.0.0';

  return template
    .evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 引入 HTML 部分片段
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doPost(e){
  const data = JSON.parse(e.postData.contents || '{}');
  const action = e.parameter.action;

  let result = {};

  switch(action){
    case 'createRoom':
      result = createRoom(data.name, data.avatarData);
      break;
    case 'joinRoom':
      result = joinRoom(data.roomId, data.name, data.avatarData);
      break;
    case 'postChat':
      result = postChat(data.roomId, data.playerId, data.text);
      break;
    // 其他動作...
    default:
      result = { error: '無效動作' };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 房間操作 =====
/**
 * 建立房間
 * @param {string} playId 玩家帳號 ID（必填，已登入）
 * @param {string} avatarUrl 玩家頭像 URL（可選）
 * @param {string} customRoomId 自訂房號（可選）
 */
function createRoom(playId, avatarUrl, customRoomId){
  if(!playId) return { error:'尚未登入' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const rooms = _loadRooms();

    // 生成房號
    let rid = customRoomId
      ? customRoomId.toUpperCase()
      : 'R' + Math.random().toString(36).slice(2,7).toUpperCase();

    if(rooms[rid]) return { error:'房號已存在' };

    // 防止同帳號重複進房
    for(const r of Object.values(rooms)){
      if(Object.values(r.players||{}).some(p=>p.accountId===playId)){
        return { error:'此帳號已在其他房間中' };
      }
    }

    // 取得玩家資訊
    const playerInfo = getPlayerByPlayId(playId);
    if(!playerInfo) return { error:'帳號不存在' };

    const pid = _uid('P');

    rooms[rid] = {
      id: rid,
      hostId: pid,
      createdAt: _now(),
      lastActive: Date.now(),
      phase: 'lobby',
      round: 0,
      players: {},
      chat: [],
      night: {},
      votes: {}
    };

    rooms[rid].players[pid] = {
      id: pid,
      accountId: playId,      // 帳號綁定
      name: playerInfo.name,   // 使用帳號名稱，不信任前端輸入
      avatar: avatarUrl || '',
      alive: true,
      role: null,
      joinedAt: _now()
    };

    rooms[rid].chat.push({
      time: _now(),
      system: true,
      text: `${rooms[rid].players[pid].name} 建立了房間`
    });

    _saveRooms(rooms);

    return { roomId: rid, playerId: pid };
  } finally {
    lock.releaseLock();
  }
}


function joinRoom(roomId, playId, avatarUrl){
  if(!playId) return { error:'尚未登入' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const rooms = _loadRooms();
    if(!rooms[roomId]) return { error:'房間不存在' };

    const room = rooms[roomId];

    // 🔒 防止同帳號重複進房（或多房）
    for(const r of Object.values(rooms)){
      if(Object.values(r.players||{}).some(p=>p.accountId===playId)){
        return { error:'此帳號已在其他房間中' };
      }
    }

    // 🔎 從帳號表取得名稱
    const playerInfo = getPlayerByPlayId(playId);
    if(!playerInfo) return { error:'帳號不存在' };

    const baseName = String(playerInfo.name || '玩家').trim();

    // 名稱去重（同房間）
    let finalName = baseName;
    const existingNames = Object.values(room.players||{}).map(p=>p.name);
    let counter = 1;
    while(existingNames.includes(finalName)){
      finalName = `${baseName}(${counter++})`;
    }

    const pid = _uid('P');

    room.players[pid] = {
      id: pid,
      accountId: playId,        // ⭐ 帳號綁定
      name: finalName,
      avatar: avatarUrl || '',
      alive: true,
      role: null,
      joinedAt: _now()
    };

    room.chat.push({
      time:_now(),
      system:true,
      text:`${finalName} 已加入房間`
    });

    _touchRoom(room);
    _saveRooms(rooms);

    return { playerId: pid, room:_stripForClient(room,pid) };
  } finally {
    lock.releaseLock();
  }
}

function leaveRoom(roomId, playerId){
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const rooms = _loadRooms();
    if(!rooms[roomId]) return;

    const room = rooms[roomId];
    if(!room.players || !room.players[playerId]) return;

    const name = room.players[playerId].name;
    delete room.players[playerId];

    room.chat.push({
      time:_now(),
      system:true,
      text:`${name} 離開房間`
    });

    _touchRoom(room);
    _saveRooms(rooms);
  } finally {
    lock.releaseLock();
  }
}


function uploadAvatar(dataUrl,filename){
  if(!dataUrl) return '';
  const matches=dataUrl.match(/^data:(.+);base64,(.+)$/); if(!matches) return '';
  const contentType=matches[1], base64=matches[2];
  const bytes=Utilities.base64Decode(base64);
  const folder=_getAvatarFolder();
  const cleanName=(filename||'avatar')+'_'+Math.floor(Math.random()*10000);
  const blob=Utilities.newBlob(bytes,contentType,cleanName);
  const file=folder.createFile(blob);
  try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW); }catch(e){}
  return file.getDownloadUrl?file.getDownloadUrl():file.getUrl();
}

// ===== 房間狀態 =====
function getRoomState(roomId, requesterId){
  const rooms=_loadRooms(); if(!rooms[roomId]) return { error:'房間不存在' };
  return _stripForClient(rooms[roomId],requesterId);
}

function _stripForClient(room, requesterId){
  const r={ id:room.id, hostId:room.hostId, phase:room.phase, round:room.round, chat:room.chat||[], players:{}, votes:room.votes||{} };
  Object.values(room.players||{}).forEach(p=>{
    let showRole=false;
    if(room.phase==='ended') showRole=true;
    else if(p.id===requesterId) showRole=true;
    r.players[p.id]={ id:p.id, name:p.name, avatar:p.avatar||'', alive:!!p.alive, role: showRole?p.role:null };
  });
  if(requesterId && room.night && room.night.checks && room.night.checks[requesterId]) r.myCheck=room.night.checks[requesterId];
  return r;
}

// ===== 分配身分 =====
function assignRoles(roomId,callerId){
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const rooms=_loadRooms(); const room=rooms[roomId]; if(!room) return {error:'房間不存在'};
    if(room.hostId!==callerId) return {error:'只有房主能分配身分'};
    const players=Object.values(room.players||{}); if(players.length<4) return {error:'玩家不足（至少 4 人）'};
    let roles=[];
    if(players.length===CONFIG.DEFAULT_PLAYERS) roles=CONFIG.ROLE_DISTRIBUTION.slice();
    else{
      let wolves=Math.max(1,Math.floor(players.length/3));
      roles=[]; for(let i=0;i<wolves;i++) roles.push('werewolf');
      roles.push('seer'); roles.push('doctor');
      while(roles.length<players.length) roles.push('villager');
    }
    roles=_shuffle(roles);
    const shuffledPlayers=_shuffle(players.map(p=>p.id));
    shuffledPlayers.forEach((pid,i)=>{ room.players[pid].role=roles[i]; room.players[pid].alive=true; });
    room.phase='rolesAssigned'; room.round=0; room.night={kills:{},saves:{},checks:{}}; room.chat.push({time:_now(),system:true,text:'房主分配了身分'});
    _touchRoom(room);_saveRooms(rooms); return {ok:true};
  } finally{ lock.releaseLock(); }
}

// ===== 夜晚操作 =====
function submitNightAction(roomId, playerId, action){
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const rooms = _loadRooms();
    const room = rooms[roomId]; if(!room) return { error:'room not found' };
    const p = room.players[playerId]; if(!p || !p.alive) return { error:'player invalid' };

    let actionText = '';
    if(action.type==='kill' && p.role==='werewolf'){
      room.night.kills[playerId] = action.targetId;
      actionText = `狼人已選擇攻擊目標`;
    } else if(action.type==='check' && p.role==='seer'){
      room.night.checks[playerId] = {targetId:action.targetId, role: room.players[action.targetId]?.role || null};
      actionText = `預言家已完成查驗`;
    } else if(action.type==='save' && p.role==='doctor'){
      room.night.saves[playerId] = action.targetId;
      actionText = `醫生已完成守護`;
    } else return { error:'action not permitted' };

    // 發布系統訊息到聊天室
    room.chat.push({time:_now(), system:true, text: actionText});

    _touchRoom(room);
    _saveRooms(rooms);
    return { ok:true };
  } finally { lock.releaseLock(); }
}

// ===== 自動結束夜晚 =====
function resolveNight(roomId, callerId){
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const rooms=_loadRooms(); const room=rooms[roomId]; if(!room) return {error:'room not found'};
    if(room.hostId!==callerId) return {error:'only host can resolve night'};
    const kills=room.night.kills||{}, saves=room.night.saves||{}; const tally={};
    Object.values(kills).forEach(tid=>{ tally[tid]=(tally[tid]||0)+1; });
    let targetId=null; if(Object.keys(tally).length>0) targetId=Object.entries(tally).sort((a,b)=>b[1]-a[1])[0][0];
    let saved=false; if(targetId && Object.values(saves||{}).indexOf(targetId)!==-1) saved=true;
    let killedName=null;
    if(targetId && !saved && room.players[targetId]){ room.players[targetId].alive=false; killedName=room.players[targetId].name; room.chat.push({time:_now(),system:true,text:`夜晚結果：${killedName} 死亡`}); }
    else if(targetId && saved) room.chat.push({time:_now(),system:true,text:`夜晚結果：有人被攻擊，但被守衛救下`});
    else room.chat.push({time:_now(),system:true,text:`夜晚結果：沒有攻擊發生`});
    room.night={kills:{},saves:{},checks:{}};
    room.phase='day'; room.round=(room.round||0)+1; room.votes={};
    _evaluateWin(room); _touchRoom(room); _saveRooms(rooms); return {ok:true, killed:killedName||null};
  } finally{ lock.releaseLock(); }
}

// ===== 投票 =====
function submitVote(roomId,voterId,targetId){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const rooms=_loadRooms(); const room=rooms[roomId]; if(!room) return {error:'room not found'};
    room.votes=room.votes||{}; room.votes[voterId]=targetId; _saveRooms(rooms); return {ok:true};
  } finally{ lock.releaseLock(); }
}

function resolveVotes(roomId,callerId){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const rooms=_loadRooms(); const room=rooms[roomId]; if(!room) return {error:'room not found'};
    if(room.hostId!==callerId) return {error:'only host can resolve votes'};
    const votes=room.votes||{}; const tally={};
    Object.values(votes).forEach(tgt=>{ if(tgt) tally[tgt]=(tally[tgt]||0)+1; });
    if(Object.keys(tally).length>0){
      const victimId=Object.entries(tally).sort((a,b)=>b[1]-a[1])[0][0];
      if(room.players[victimId]){ room.players[victimId].alive=false; room.chat.push({time:_now(),system:true,text:`投票結果：${room.players[victimId].name} 被處決`}); }
    } else room.chat.push({time:_now(),system:true,text:`投票無效：沒有人被投票`});
    room.votes={}; _evaluateWin(room); if(room.phase!=='ended') room.phase='night'; _touchRoom(room); _saveRooms(rooms); return {ok:true};
  } finally{ lock.releaseLock(); }
}

function _evaluateWin(room){
  const players = Object.values(room.players||{});
  const wolvesAlive = players.filter(p=>p.alive && p.role==='werewolf').length;
  const othersAlive = players.filter(p=>p.alive && p.role!=='werewolf').length;

  if(wolvesAlive===0){
    room.phase='ended';
    room.winner='villagers';
    room.chat.push({time:_now(),system:true,text:'遊戲結束：村民勝利'});
  }
  else if(wolvesAlive>=othersAlive){
    room.phase='ended';
    room.winner='werewolves';
    room.chat.push({time:_now(),system:true,text:'遊戲結束：狼人勝利'});
  }

  if(room.phase==='ended'){
    players.forEach(p=>{
      const isWin =
        (room.winner==='villagers' && p.role!=='werewolf') ||
        (room.winner==='werewolves' && p.role==='werewolf');

      _updatePlayerStats(p.accountId, isWin); // ⭐ 正確
    });
  }
}

function getPlayerByPlayId(playId){
  const sheet = _getOrCreatePlayerSheet();
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]) === String(playId)){
      return {
        playId: data[i][0],
        name: data[i][1],
        wins: Number(data[i][3])||0,
        losses: Number(data[i][4])||0,
        winRate: Number(data[i][5])||0
      };
    }
  }
  return null;
}


function postChat(roomId, playerId, text){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const rooms=_loadRooms(); if(!rooms[roomId]) return {error:'room not found'};
    const room=rooms[roomId]; const name=room.players[playerId]?room.players[playerId].name:'匿名';
    room.chat.push({time:_now(),name:name,text:text}); _touchRoom(room); _saveRooms(rooms); return {ok:true};
  } finally{ lock.releaseLock(); }
}

// 清理空房間（沒有玩家）
function cleanupEmptyRooms() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rooms = _loadRooms();
    let changed = false;
    Object.keys(rooms).forEach(rid => {
      if (!rooms[rid].players || Object.keys(rooms[rid].players).length === 0) {
        delete rooms[rid];
        changed = true;
      }
    });
    if (changed) _saveRooms(rooms);
  } finally {
    lock.releaseLock();
  }
}

// 取得所有房間列表（簡化資料給前端）
function listRooms() {
  cleanupEmptyRooms();
  cleanupInactiveRooms();
  const rooms = _loadRooms();
  const now = Date.now();
  return Object.values(rooms).map(r => {
    const lastActive = r.lastActive ? new Date(r.lastActive).getTime() : new Date(r.createdAt).getTime();
    const inactive = (now - lastActive) > 60 * 60 * 1000;
    return {
      id: r.id,
      hostName: r.players[r.hostId]?.name || '房主',
      playerCount: Object.keys(r.players||{}).length,
      lastActive: lastActive,
      inactive: inactive
    };
  });
}

function loginPlayer(name, password) {
  const sheet = _getOrCreatePlayerSheet();
  const data = sheet.getDataRange().getValues();

  name = String(name).trim();
  password = String(password).trim();

  for (let i = 1; i < data.length; i++) {
    const sheetName = String(data[i][1]).trim();
    const sheetPass = String(data[i][2]).trim();

    if (sheetName === name && sheetPass === password) {
      return {
        playId: String(data[i][0]),
        name: sheetName,
        wins: Number(data[i][3]) || 0,
        losses: Number(data[i][4]) || 0,
        winRate: Number(data[i][5]) || 0
      };
    }
  }
  return { error: '帳號或密碼錯誤' };
}


function registerPlayer(name, password) {
  const sheet = _getOrCreatePlayerSheet();
  const data = sheet.getDataRange().getValues();
  
  // 名稱不能重複
  for(let i=1;i<data.length;i++){
    if(data[i][1]===name) return {error:'名稱已被使用'};
  }
  
  const playId = _uid('P');
  const wins = 0, losses = 0;
  const winRate = 0;
  
  sheet.appendRow([playId, name, password, wins, losses, winRate]);
  return {ok:true, playId};
}


function _getOrCreatePlayerSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName('Players');
  if(!sheet) {
    sheet = ss.insertSheet('Players');
    sheet.getRange(1,1,1,6).setValues([['playId','name','password','wins','losses','winRate']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _updatePlayerStats(playerId, isWin){
  const sheet = _getOrCreatePlayerSheet();
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(data[i][0]===playerId){
      let wins = parseInt(data[i][3]);
      let losses = parseInt(data[i][4]);
      if(isWin) wins++; else losses++;
      const winRate = (wins + losses) ? Math.round(wins / (wins + losses) * 100) : 0;
      sheet.getRange(i+1, 4, 1, 3).setValues([[wins, losses, winRate]]);
      break;
    }
  }
}


// 驗證密碼
function adminLogin(password) {
  return password === ADMIN_PASSWORD;
}

// 取得房間列表 (簡化資料)
function adminListRooms(password) {
  if(!adminLogin(password)) return {error:"密碼錯誤"};
  const rooms = _loadRooms();
  return Object.values(rooms).map(r=>({
    id: r.id,
    hostName: r.players[r.hostId]?.name || '房主',
    playerCount: Object.keys(r.players||{}).length,
    lastActive: r.lastActive,
    phase: r.phase
  }));
}

function adminDeleteRoom(roomId){
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const rooms = _loadRooms();
    if(!rooms[roomId]) return {error:"房間不存在"};
    delete rooms[roomId];
    _saveRooms(rooms);
    return {ok:true,message:`房間 ${roomId} 已刪除`};
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteAllRooms(){
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    _saveRooms({});
    return {ok:true,message:"所有房間已清空"};
  } finally {
    lock.releaseLock();
  }
}


// 自動清理空房間或長時間無活動的房間
function cleanupInactiveRooms() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 最多等待 30 秒
  try {
    const rooms = _loadRooms();
    const now = Date.now();
    let changed = false;

    Object.keys(rooms).forEach(rid => {
      const room = rooms[rid];
      const lastActive = room.lastActive ? new Date(room.lastActive).getTime() : new Date(room.createdAt).getTime();

      // 條件：房間無玩家 或 最後活動超過 60 分鐘
      const noPlayers = !room.players || Object.keys(room.players).length === 0;
      const inactive = (now - lastActive) > 60 * 60 * 1000; // 60 分鐘
      if (noPlayers || inactive) {
        delete rooms[rid];
        changed = true;
        Logger.log(`刪除房間 ${rid} (無人或超過 60 分鐘無活動)`);
      }
    });

    if (changed) _saveRooms(rooms);
  } finally {
    lock.releaseLock();
  }
}

function createCleanupTrigger() {
  // 先刪掉已存在的相同觸發器，避免重複
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'cleanupInactiveRooms') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 建立新的時間觸發器，每 10 分鐘執行一次
  ScriptApp.newTrigger('cleanupInactiveRooms')
           .timeBased()
           .everyMinutes(10)
           .create();
}

function _touchRoom(room) {
  // 更新最後活動時間，用來做自動清除閒置房間
  room.lastActive = Date.now();
}

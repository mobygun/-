/* 로그인 + 클라우드 동기화 (Firebase). 로그인하지 않으면 지금처럼 폰에만 저장됩니다. */
(function(global){
  const CFG = {
    apiKey: "AIzaSyBci4ypjpgYeQdjgwa0Ta_5yE4whWaRdFA",
    authDomain: "tscore-app-c4a10.firebaseapp.com",
    projectId: "tscore-app-c4a10",
    storageBucket: "tscore-app-c4a10.firebasestorage.app",
    messagingSenderId: "689202684050",
    appId: "1:689202684050:web:2aecbdb075600b32fec64b"
  };
  const DOMAIN = "@tscore.kr";           // 아이디 뒤에 자동으로 붙는 주소
  const SDK = "https://www.gstatic.com/firebasejs/10.12.2/";

  const S = {
    ready:false, user:null, isAdmin:false, unsub:null,
    viewUid:null,          // 어드민이 다른 직원 기록을 볼 때
    onState:null,          // 상태가 바뀌면 앱에 알림
    onRemote:null          // 서버 데이터가 도착하면 앱에 전달
  };
  global.SYNC = S;

  function loadScript(src){
    return new Promise((ok,ng)=>{
      const el=document.createElement("script");
      el.src=src; el.onload=ok; el.onerror=()=>ng(new Error(src));
      document.head.appendChild(el);
    });
  }

  S.init = async function(){
    if(S.ready) return true;
    try{
      await loadScript(SDK+"firebase-app-compat.js");
      await loadScript(SDK+"firebase-auth-compat.js");
      await loadScript(SDK+"firebase-firestore-compat.js");
      firebase.initializeApp(CFG);
      S.auth = firebase.auth();
      S.db = firebase.firestore();
      try{ await S.db.enablePersistence({synchronizeTabs:true}); }catch(e){}
      await S.auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL);
      S.auth.onAuthStateChanged(async (u)=>{
        S.user=u; S.isAdmin=false; S.viewUid=null;
        if(S.unsub){ S.unsub(); S.unsub=null; }
        if(u){
          try{
            const a=await S.db.collection("admins").doc(u.uid).get();
            S.isAdmin=a.exists;
          }catch(e){}
          S.watch(u.uid);
        }
        if(S.onState) S.onState();
      });
      S.ready=true;
      return true;
    }catch(e){
      console.error(e);
      return false;
    }
  };

  function idToEmail(id){
    id=String(id||"").trim();
    return id.includes("@") ? id : id+DOMAIN;
  }

  S.signup = async function(id, pw, name){
    await S.init();
    const cred = await S.auth.createUserWithEmailAndPassword(idToEmail(id), pw);
    await S.db.collection("users").doc(cred.user.uid).set({
      id:String(id).trim(), name:name||String(id).trim(), updatedAt:Date.now()
    },{merge:true});
    return cred.user;
  };
  S.login = async function(id, pw){
    await S.init();
    const cred = await S.auth.signInWithEmailAndPassword(idToEmail(id), pw);
    return cred.user;
  };
  S.logout = async function(){
    if(!S.ready) return;
    if(S.unsub){ S.unsub(); S.unsub=null; }
    await S.auth.signOut();
  };

  // 서버 데이터 실시간 감시
  S.watch = function(uid){
    if(S.unsub){ S.unsub(); S.unsub=null; }
    S.unsub = S.db.collection("users").doc(uid).onSnapshot((doc)=>{
      if(!doc.exists) return;
      const d=doc.data();
      if(S.onRemote) S.onRemote(d, uid);
    }, (err)=>console.error(err));
  };

  // 앱 데이터 업로드
  let timer=null, pending=null;
  S.push = function(data){
    if(!S.user || S.viewUid) return;           // 로그아웃 상태거나 남의 기록 보는 중이면 저장 안 함
    pending=data;
    clearTimeout(timer);
    timer=setTimeout(async ()=>{
      try{
        await S.db.collection("users").doc(S.user.uid).set(
          Object.assign({updatedAt:Date.now()}, pending), {merge:true});
      }catch(e){ console.error(e); }
    }, 800);
  };

  // 본인 계정 탈퇴 (본인 것만 가능)
  S.deleteAccount = async function(pw){
    if(!S.user) throw new Error("로그인이 필요합니다.");
    const cred = firebase.auth.EmailAuthProvider.credential(S.user.email, pw);
    await S.user.reauthenticateWithCredential(cred);
    try{ await S.db.collection("users").doc(S.user.uid).delete(); }catch(e){}
    await S.user.delete();
  };

  // 주간업무일지: 공용 문서 실시간 감시
  let weeklyUnsub=null;
  S.watchWeekly = function(weekId){
    if(weeklyUnsub){ weeklyUnsub(); weeklyUnsub=null; }
    weeklyUnsub = S.db.collection("weeklyLogs").doc(weekId).onSnapshot((doc)=>{
      if(S.onWeeklyRemote) S.onWeeklyRemote(doc.exists ? doc.data() : null, weekId);
    }, (err)=>console.error(err));
  };
  S.stopWeekly = function(){ if(weeklyUnsub){ weeklyUnsub(); weeklyUnsub=null; } };

  // 주간업무일지: 공용 문서 저장 (로그인한 사람이면 누구나 전체 문서를 씀)
  let weeklyTimer=null, weeklyPending=null, weeklyPendingId=null;
  S.pushWeekly = function(weekId, data){
    if(!S.user) return;
    weeklyPendingId = weekId; weeklyPending = data;
    clearTimeout(weeklyTimer);
    weeklyTimer = setTimeout(async ()=>{
      try{
        await S.db.collection("weeklyLogs").doc(weeklyPendingId).set(
          Object.assign({updatedAt:Date.now(), updatedBy:S.user.uid}, weeklyPending), {merge:true});
      }catch(e){ console.error(e); }
    }, 800);
  };

  // 주간업무일지: 주차 목록 (weekId 내림차순)
  S.listWeeks = async function(){
    const snap = await S.db.collection("weeklyLogs").orderBy("weekStart","desc").get();
    return snap.docs.map(d=>({weekId:d.id, ...d.data()}));
  };

  // 어드민: 전체 사용자 목록
  S.listUsers = async function(){
    const snap = await S.db.collection("users").get();
    return snap.docs.map(d=>({uid:d.id, id:d.data().id||"", name:d.data().name||d.id}));
  };
  // 어드민: 특정 사용자 기록 보기
  S.viewUser = async function(uid){
    S.viewUid = (uid===S.user.uid) ? null : uid;
    const doc = await S.db.collection("users").doc(uid).get();
    if(doc.exists && S.onRemote) S.onRemote(doc.data(), uid);
  };
})(window);

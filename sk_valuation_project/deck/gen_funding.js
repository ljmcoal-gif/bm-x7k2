// SK Funding Structure Deck — 재무팀 반영 v2 펀딩 구조 (4슬라이드)
// 디자인: gen_v2.js와 동일한 SK 실적발표 템플릿 (A4 가로)
//   배경 #FBF9F5 / 헤드라인 Noto Serif KR Bold / 본문 IBM Plex Sans KR
//   수치 IBM Plex Mono / 레드 #C1121F / 웜그레이 #D8CFC2 / 그린 #1E7A4D
//   그라데이션·3D·그림자 금지 · 가로 그리드만
// 슬라이드: 1) 요약  2) 법인별 자금구조  3) 이렇게 설정한 이유  4) 정부와의 Comm
const pptxgen = require(require('child_process').execSync('npm root -g').toString().trim() + '/pptxgenjs');
const pres = new pptxgen();

// ===== A4 가로 =====
const W = 11.69, H = 8.27;
pres.defineLayout({ name: 'A4L', width: W, height: H });
pres.layout = 'A4L';

// ===== 디자인 토큰 (gen_v2.js와 동일) =====
const BG="FBF9F5", INK="1B1714", DARK="1B1714", RED="C1121F", WARM="D8CFC2",
      GREEN="1E7A4D", GRID="ECE6DC", RULE="E4DED4", MUTE="6F6960", MUTE2="9A847C", PAPER="FFFFFF";
const HEAD="Noto Serif CJK KR", BODY="IBM Plex Sans KR", MONO="IBM Plex Mono";
const MX=0.62, TOP=0.52;

// ===== 공통 헬퍼 (gen_v2.js와 동일) =====
function base(s){ s.background = { color: BG }; }
function header(s, eyebrow, page, title, dek){
  s.addText(eyebrow, { x:MX, y:TOP, w:7.5, h:0.26, margin:0, align:"left",
    fontFace:MONO, fontSize:9, color:RED, charSpacing:2, valign:"middle" });
  s.addText("SK Innovation · IR  /  "+page, { x:W-4.5-MX, y:TOP, w:4.5, h:0.26, margin:0, align:"right",
    fontFace:MONO, fontSize:9, color:MUTE2, charSpacing:1.5, valign:"middle" });
  s.addText(title, { x:MX, y:TOP+0.34, w:W-2*MX-2.4, h:0.62, margin:0, align:"left",
    fontFace:HEAD, fontSize:24, bold:true, color:INK, valign:"middle" });
  if(dek){
    s.addText(dek, { x:W-MX-3.0, y:TOP+0.5, w:3.0, h:0.42, margin:0, align:"right",
      fontFace:BODY, fontSize:9.5, color:MUTE, valign:"bottom" });
  }
  s.addShape(pres.shapes.LINE, { x:MX, y:TOP+1.04, w:W-2*MX, h:0, line:{ color:INK, width:1.5 } });
}
function footer(s){
  s.addShape(pres.shapes.LINE, { x:MX, y:H-0.46, w:W-2*MX, h:0, line:{ color:RULE, width:0.75 } });
  s.addText("Project [ ● ]  ·  Strictly Private & Confidential", { x:MX, y:H-0.42, w:5, h:0.26, margin:0,
    fontFace:BODY, fontSize:7.5, color:MUTE2, valign:"middle" });
  s.addText("DRAFT — FOR DISCUSSION ONLY  ·  ★ = 미검증(약정서·법률 확인 필요)", { x:W/2-3, y:H-0.42, w:6, h:0.26, margin:0, align:"center",
    fontFace:MONO, fontSize:7, color:MUTE2, charSpacing:1, valign:"middle" });
  s.addText("2026", { x:W-MX-1.5, y:H-0.42, w:1.5, h:0.26, margin:0, align:"right",
    fontFace:MONO, fontSize:7.5, color:MUTE2, valign:"middle" });
}
function eyebrowLabel(s, x, y, w, txt){
  s.addText(txt, { x, y, w, h:0.24, margin:0, fontFace:MONO, fontSize:8.5, color:MUTE2,
    charSpacing:1.5, valign:"middle" });
}

(async ()=>{

  // ════════════ SLIDE 1 — 요약 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"FUNDING STRUCTURE v2 — EXECUTIVE SUMMARY","01",
      "펀딩 구조 v2 — 요약","재무팀 의견 반영");

    // 핵심 명제 (다크 박스)
    const tw=W-2*MX;
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:1.95,w:tw,h:1.02,fill:{color:DARK},line:{type:"none"}});
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:1.95,w:0.07,h:1.02,fill:{color:RED}});
    s.addText("핵심 명제",{x:MX+0.25,y:2.06,w:tw-0.5,h:0.24,margin:0,fontFace:MONO,fontSize:8.5,color:WARM,charSpacing:1.5,valign:"middle"});
    s.addText([{text:"정부 앵커를 OT→BA로 이동",options:{bold:true,color:PAPER}},
      {text:"하되, BA는 ",options:{color:WARM}},
      {text:"보조금(grant) 우선",options:{bold:true,color:"FF9B9B"}},
      {text:" · Equity는 잔여·최후로 둔다. ",options:{color:WARM}},
      {text:"OT의 DOE 링펜스는 유지",options:{bold:true,color:"9BD3A8"}},
      {text:".",options:{color:WARM}}],
      {x:MX+0.25,y:2.32,w:tw-0.5,h:0.55,margin:0,fontFace:HEAD,fontSize:15,valign:"middle",lineSpacingMultiple:1.0});

    // 전환 포인트 3
    eyebrowLabel(s,MX,3.2,tw,"v1 → v2  전환 포인트");
    const cards=[
      ["정부자금 타깃","OT → BA","OT는 ATVM 이미 소진 → 신규 미정부자금은 BA측 feasibility ↑"],
      ["BA 조달 방식","보조금 우선","무상환·무DSCR·무희석·CoC 무관 → EBIT=0 자산에 유일하게 깨끗"],
      ["OT 구조","링펜스 유지","DOE $4bn 선순위·CoC·5층 격벽 그대로 → SKI 보증을 전염에서 차단"],
    ];
    const cw=(tw-0.6)/3; let cx=MX;
    cards.forEach((c,i)=>{
      s.addShape(pres.shapes.RECTANGLE,{x:cx,y:3.48,w:cw,h:1.5,fill:{color:PAPER},line:{color:RULE,width:1}});
      s.addShape(pres.shapes.RECTANGLE,{x:cx,y:3.48,w:cw,h:0.05,fill:{color:RED}});
      s.addText(c[0],{x:cx+0.18,y:3.62,w:cw-0.36,h:0.26,margin:0,fontFace:MONO,fontSize:8,color:MUTE2,charSpacing:1,valign:"middle"});
      s.addText(c[1],{x:cx+0.18,y:3.88,w:cw-0.36,h:0.42,margin:0,fontFace:HEAD,bold:true,fontSize:17,color:RED,valign:"middle"});
      s.addText(c[2],{x:cx+0.18,y:4.34,w:cw-0.36,h:0.6,margin:0,fontFace:BODY,fontSize:9,color:MUTE,valign:"top",lineSpacingMultiple:1.05});
      cx+=cw+0.3;
    });

    // BA 펀딩 워터폴
    eyebrowLabel(s,MX,5.25,tw,"BA 펀딩 워터폴 — 조달 우선순위");
    const wf=[["①","보조금 (1순위)","MESC·48C·州 인센티브 — 최대화 ★",RED],
      ["②","비소구 자산부채 (제한)","설비·AMPC 수취권 담보 · SKI 보증 미가산",INK],
      ["③","Equity (잔여·최후)","CoC 20% 캡 내 · DOE Equity 불가 가정 ★",MUTE]];
    const ww=(tw-1.0)/3; let wx=MX;
    wf.forEach((p,i)=>{
      s.addShape(pres.shapes.RECTANGLE,{x:wx,y:5.55,w:ww,h:1.05,fill:{color: i===0?"FBF3F2":PAPER},line:{color: i===0?RED:RULE,width: i===0?1.5:1}});
      s.addText(p[0],{x:wx+0.15,y:5.66,w:0.5,h:0.4,margin:0,fontFace:HEAD,bold:true,fontSize:18,color:p[3],valign:"middle"});
      s.addText(p[1],{x:wx+0.62,y:5.66,w:ww-0.75,h:0.4,margin:0,fontFace:BODY,bold:true,fontSize:11,color:INK,valign:"middle"});
      s.addText(p[2],{x:wx+0.15,y:6.08,w:ww-0.3,h:0.45,margin:0,fontFace:BODY,fontSize:8.5,color:MUTE,valign:"top",lineSpacingMultiple:1.05});
      if(i<2) s.addText("▶",{x:wx+ww-0.02,y:5.55,w:0.5,h:1.05,margin:0,align:"center",fontFace:BODY,fontSize:12,color:WARM,valign:"middle"});
      wx+=ww+0.5;
    });
    s.addText("상위 단계로 충당 가능한 금액은 하위로 내리지 않는다 — 희석·부채부담 최소화.",
      {x:MX,y:6.78,w:tw,h:0.3,margin:0,fontFace:BODY,fontSize:9,color:INK,valign:"middle"});
    footer(s);
  }

  // ════════════ SLIDE 2 — 법인별 자금 구조 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"FUNDING BY ENTITY — STRUCTURE MAP","02",
      "법인별 자금 구조 — BA · OT 분리","HoldCo 우산 + 격벽");

    const tw=W-2*MX;
    // HoldCo 우산 (상단 바)
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:1.9,w:tw,h:0.62,fill:{color:DARK},line:{type:"none"}});
    s.addText([{text:"HoldCo 우산  ",options:{bold:true,color:PAPER,fontSize:12}},
      {text:"통합조달(협상력↑·조달비용↓) — 위에서 묶고, 아래(SPV)에서 격리한다  ·  하이브리드(B) 구조",options:{color:WARM,fontSize:9.5}}],
      {x:MX+0.25,y:1.9,w:tw-0.5,h:0.62,margin:0,fontFace:BODY,valign:"middle"});

    // 두 법인 카드
    const colW=(tw-0.4)/2; const cardY=2.72, cardH=3.62;
    const cards=[
      {x:MX, tint:INK, name:"공장 1 — SKBA (조지아)", sub:"파우치 21.2GWh · DC블록 개조 대상",
       use:"용도 = BA 개조(retrofit) 자금",
       rows:[
         ["①","보조금 (1순위)","MESC / 48C / 州 — 무상환·무희석·무CoC ★",RED],
         ["②","비소구 자산부채","설비·재고·AMPC 수취권 담보 · 한도 내",INK],
         ["③","Equity (잔여)","CoC 20% 캡 내 · 밸류 플로어·anti-dilution",MUTE],
       ],
       warn:"제약: 기존 차입 ~$5bn + SK이노 보증·확약 → 신규 대출 회피, 레버리지/covenant 관리 ★"},
      {x:MX+colW+0.4, tint:RED, name:"공장 2 — SKOT (테네시)", sub:"각형+파우치 29.7GWh · DOE 묶임",
       use:"용도 = 기존 구조 유지 + AMPC 창 내 회수",
       rows:[
         ["•","DOE $4bn ATVM","기존 선순위 담보 — 이자만(원금 0)",INK],
         ["•","AMPC 현금화","45X 수취권 유동화 → 미래현금 당겨옴",INK],
         ["•","메자닌 $1bn @7%","보조금 창(2029~32) 내 상환 · DSCR 1.3x",INK],
       ],
       warn:"링펜스 5층 유지: 법인분리·Separateness·Intercreditor·담보·Non-recourse"},
    ];
    cards.forEach(c=>{
      s.addShape(pres.shapes.RECTANGLE,{x:c.x,y:cardY,w:colW,h:cardH,fill:{color:PAPER},line:{color:RULE,width:1}});
      // 타이틀 바
      s.addShape(pres.shapes.RECTANGLE,{x:c.x,y:cardY,w:colW,h:0.5,fill:{color:c.tint},line:{type:"none"}});
      s.addText(c.name,{x:c.x+0.18,y:cardY+0.04,w:colW-0.36,h:0.3,margin:0,fontFace:BODY,bold:true,fontSize:12,color:PAPER,valign:"middle"});
      s.addText(c.sub,{x:c.x+0.18,y:cardY+0.32,w:colW-0.36,h:0.18,margin:0,fontFace:MONO,fontSize:7.5,color:WARM,valign:"middle"});
      // 용도
      s.addText(c.use,{x:c.x+0.18,y:cardY+0.6,w:colW-0.36,h:0.26,margin:0,fontFace:BODY,bold:true,fontSize:9.5,color:INK,valign:"middle"});
      // rows
      let ry=cardY+0.95;
      c.rows.forEach(r=>{
        s.addShape(pres.shapes.RECTANGLE,{x:c.x+0.18,y:ry,w:colW-0.36,h:0.62,fill:{color:BG},line:{color:GRID,width:0.5}});
        s.addText(r[0],{x:c.x+0.28,y:ry,w:0.4,h:0.62,margin:0,align:"center",fontFace:HEAD,bold:true,fontSize:15,color:r[3],valign:"middle"});
        s.addText(r[1],{x:c.x+0.72,y:ry+0.06,w:colW-0.9,h:0.28,margin:0,fontFace:BODY,bold:true,fontSize:10.5,color:INK,valign:"middle"});
        s.addText(r[2],{x:c.x+0.72,y:ry+0.32,w:colW-0.9,h:0.26,margin:0,fontFace:BODY,fontSize:8.5,color:MUTE,valign:"middle"});
        ry+=0.68;
      });
      // warn 박스
      s.addShape(pres.shapes.RECTANGLE,{x:c.x+0.18,y:cardY+cardH-0.66,w:colW-0.36,h:0.52,fill:{color:"F2ECE0"},line:{type:"none"}});
      s.addText(c.warn,{x:c.x+0.3,y:cardY+cardH-0.66,w:colW-0.6,h:0.52,margin:0,fontFace:BODY,fontSize:8,color:INK,valign:"middle",lineSpacingMultiple:1.0});
    });

    // Non-recourse 하단 바
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:6.5,w:tw,h:0.55,fill:{color:DARK},line:{type:"none"}});
    s.addText([{text:"Non-recourse 격벽  ",options:{bold:true,color:"9BD3A8",fontSize:11}},
      {text:"각 SPV 채무는 자기 자산에만 청구 — 모회사·타 공장·SK이노 보증에 무소구. 한 공장이 무너져도 다른 칸으로 전염되지 않는다.",options:{color:WARM,fontSize:9.5}}],
      {x:MX+0.25,y:6.5,w:tw-0.5,h:0.55,margin:0,fontFace:BODY,valign:"middle"});
    footer(s);
  }

  // ════════════ SLIDE 3 — 이렇게 설정한 이유 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"RATIONALE — WHY THIS STRUCTURE","03",
      "이렇게 설정한 이유","5가지 근거");
    const tw=W-2*MX;
    const reasons=[
      ["01","Feasibility — OT는 한 우물, BA는 새 우물","OT는 ATVM 수혜를 이미 받고 있어 신규 연방자금 여지가 작다. DOE/DOW 관점에서 BA측 funding feasibility가 더 높다."],
      ["02","보조금은 EBIT=0 자산에 유일하게 '깨끗한' 돈","DC블록은 박리다매(EBIT≈0). 상업대출은 DSCR이 안 나와 원천적으로 안 붙는다. 보조금은 무상환·무DSCR·무희석·CoC 무관 → 구조적 최적."],
      ["03","BA 대출 회피 — 이미 $5bn + SKI 보증","BA는 기존 차입이 크고 SK이노 보증·확약이 걸려 있다. 추가 대출은 레버리지 한도·cross-default·보증 노출을 키운다 → 보조금/비소구로 우회."],
      ["04","Equity는 잔여·최후 — DOE Equity는 불가 가정","BA 밸류에이션·지분희석·CoC(20% 캡)를 고려하면 Equity는 마지막. 게다가 DOE LPO/ATVM은 법령상 대출·보증이지 지분투자가 아니다 ★."],
      ["05","OT 링펜스 유지 = SKI 보증의 방패","링펜스는 비용이 아니라 보호다. OT가 2033년 이후 적자 회귀해도 Non-recourse가 BA·모회사·SKI 보증으로의 전염을 끊는다."],
    ];
    let ry=1.95; const rh=0.92;
    reasons.forEach((r,i)=>{
      const hl=(i===1||i===4);
      s.addShape(pres.shapes.RECTANGLE,{x:MX,y:ry,w:tw,h:rh-0.1,fill:{color: hl?"FBF3F2":PAPER},line:{color: hl?RED:RULE,width: hl?1.5:1}});
      s.addText(r[0],{x:MX+0.18,y:ry,w:0.7,h:rh-0.1,margin:0,align:"center",fontFace:HEAD,bold:true,fontSize:22,color: hl?RED:WARM,valign:"middle"});
      s.addText(r[1],{x:MX+1.0,y:ry+0.1,w:tw-1.2,h:0.32,margin:0,fontFace:BODY,bold:true,fontSize:12,color: hl?RED:INK,valign:"middle"});
      s.addText(r[2],{x:MX+1.0,y:ry+0.42,w:tw-1.2,h:0.4,margin:0,fontFace:BODY,fontSize:9,color:MUTE,valign:"top",lineSpacingMultiple:1.0});
      ry+=rh;
    });
    footer(s);
  }

  // ════════════ SLIDE 4 — 정부와의 Comm 방식 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"GOVERNMENT ENGAGEMENT — HOW WE COMMUNICATE","04",
      "정부와의 Comm 방식","재무 사전 정렬 우선");
    const tw=W-2*MX;
    const colW=(tw-0.5)*0.52, rx=MX+colW+0.5, rw=tw-colW-0.5;

    // 좌: 시퀀스
    eyebrowLabel(s,MX,1.95,colW,"ENGAGEMENT SEQUENCE");
    const steps=[
      ["1","내부 재무 사전 정렬 (선행·필수)","ask(grant/loan)·구조·레드라인을 재무와 먼저 합의",true],
      ["2","ask 확정 → 카운터파트 매칭","무엇을 요청할지 정한 뒤 매칭 부서로",false],
      ["3","DOE 부서 미팅","합의된 ask로만 — 부서별 단일 메시지",false],
      ["4","DOW(국방) 트랙 — 방산·딥테크 별도","드론·로보틱스 슬리브로 분리 진행",false],
    ];
    let sy=2.28; const sh=0.95;
    steps.forEach((st,i)=>{
      const hl=st[3];
      s.addShape(pres.shapes.RECTANGLE,{x:MX,y:sy,w:colW,h:sh-0.12,fill:{color: hl?DARK:PAPER},line:{color: hl?DARK:RULE,width:1}});
      s.addShape(pres.shapes.OVAL,{x:MX+0.16,y:sy+0.2,w:0.44,h:0.44,fill:{color: hl?RED:WARM},line:{type:"none"}});
      s.addText(st[0],{x:MX+0.16,y:sy+0.2,w:0.44,h:0.44,margin:0,align:"center",fontFace:HEAD,bold:true,fontSize:16,color: hl?PAPER:INK,valign:"middle"});
      s.addText(st[1],{x:MX+0.74,y:sy+0.12,w:colW-0.9,h:0.34,margin:0,fontFace:BODY,bold:true,fontSize:11,color: hl?PAPER:INK,valign:"middle"});
      s.addText(st[2],{x:MX+0.74,y:sy+0.44,w:colW-0.9,h:0.32,margin:0,fontFace:BODY,fontSize:8.5,color: hl?WARM:MUTE,valign:"top"});
      if(i<3) s.addText("▼",{x:MX+0.28,y:sy+sh-0.18,w:0.3,h:0.2,margin:0,align:"center",fontFace:BODY,fontSize:9,color:WARM,valign:"middle"});
      sy+=sh;
    });

    // 우상: 부서 매칭 표
    eyebrowLabel(s,rx,1.95,rw,"ask → DOE 카운터파트 매칭");
    const match=[["보조금 (BA)","MESC — 제조·공급망",GREEN],
      ["대출","LPO (ATVM/Title 17)  ※ OT 기존",INK],
      ["Equity","해당 부서 없음 — DOE 지분투자 ✗",RED],
      ["방산 앵글","DOW (국방) 별도 트랙",INK]];
    let my=2.28; const mh=0.5;
    match.forEach((m,i)=>{
      s.addShape(pres.shapes.RECTANGLE,{x:rx,y:my,w:rw,h:mh-0.06,fill:{color: i%2?BG:PAPER},line:{color:GRID,width:0.5}});
      s.addShape(pres.shapes.RECTANGLE,{x:rx,y:my,w:0.05,h:mh-0.06,fill:{color:m[2]}});
      s.addText(m[0],{x:rx+0.2,y:my,w:1.7,h:mh-0.06,margin:0,fontFace:BODY,bold:true,fontSize:9.5,color:INK,valign:"middle"});
      s.addText(m[1],{x:rx+1.95,y:my,w:rw-2.1,h:mh-0.06,margin:0,fontFace:BODY,fontSize:9,color:m[2],valign:"middle"});
      my+=mh;
    });

    // 우하: 레드라인 박스
    const ry2=my+0.18;
    s.addShape(pres.shapes.RECTANGLE,{x:rx,y:ry2,w:rw,h:1.72,fill:{color:DARK},line:{type:"none"}});
    s.addText("레드라인 (미팅 전 재무 합의)",{x:rx+0.2,y:ry2+0.12,w:rw-0.4,h:0.3,margin:0,fontFace:HEAD,bold:true,fontSize:12,color:PAPER,valign:"middle"});
    s.addText([
      {text:"• SK이노 보증 확대 금지\n",options:{color:WARM}},
      {text:"• BA 기존 covenant breach 금지\n",options:{color:WARM}},
      {text:"• CoC 20% 캡 유지\n",options:{color:WARM}},
      {text:"• DOE Equity 기대 제외(대출·보증만)",options:{color:"FF9B9B",bold:true}},
    ],{x:rx+0.2,y:ry2+0.5,w:rw-0.4,h:1.1,margin:0,fontFace:BODY,fontSize:10,valign:"top",lineSpacingMultiple:1.18});

    // 하단 원칙
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:H-0.95,w:tw,h:0.4,fill:{color:"F2ECE0"},line:{type:"none"}});
    s.addText([{text:"[원칙]  ",options:{bold:true,color:RED}},
      {text:"부서 미팅 전, '무엇을 요청할지'를 재무와 먼저 확정한다 — 잘못된 부서에 잘못된 ask로 가는 것을 막는다.",options:{color:INK}}],
      {x:MX+0.2,y:H-0.95,w:tw-0.4,h:0.4,margin:0,fontFace:BODY,fontSize:9.5,valign:"middle"});
    footer(s);
  }

  await pres.writeFile({ fileName: require("path").join(__dirname,"..","output","SK_Funding_v2.pptx") });
  console.log("WROTE funding deck");
})();

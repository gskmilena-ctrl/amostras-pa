<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="description" content="Controle de estoque de amostras de produto acabado">
<link rel="manifest" href="manifest.webmanifest">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Amostras PA">
<title>Amostras de Produto Acabado</title>
<meta name="theme-color" content="#121524">
<link rel="icon" type="image/png" href="icones/icone-192.png">
<link rel="apple-touch-icon" href="icones/apple-touch-icon.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.14.1/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore-compat.js"></script>
<script src="config.js"></script>
<style>
/* Layout: menu lateral com ícones (computador) e barra de navegação embaixo (celular), no estilo EFClaudian. Temas escolhidos pela pessoa (roxo escuro é o padrão). */
:root{
  --bg:#0a0a0f; --surface:#13131c; --sunk:#1a1a26; --surface3:#232332; --ink:#e8e8ed; --muted:#9a9aab; --dim:#6a6a7e; --line:#262632; --line2:#33334a;
  --accent:#8b5cf6; --accent-ink:#ffffff; --accent-soft:rgba(139,92,246,.16);
  --ok:#5bb592; --ok-soft:rgba(91,181,146,.15); --warn:#d8a657; --warn-soft:rgba(216,166,87,.15); --bad:#e07b7b; --bad-soft:rgba(224,123,123,.15); --idle:#9a9aab; --idle-soft:#232332;
  --shadow:0 14px 36px rgba(0,0,0,.5);
  --display:"Inter",-apple-system,system-ui,"Segoe UI",Roboto,sans-serif; --body:"Inter",-apple-system,system-ui,"Segoe UI",Roboto,sans-serif; --mono:"IBM Plex Mono",ui-monospace,Menlo,Consolas,monospace;
  --r:8px; --rc:12px; color-scheme:dark;
}
:root[data-skin="ouro"]{
  --bg:#0b0d16; --surface:#121524; --sunk:#181c2e; --surface3:#21263c; --ink:#f3eee3; --muted:#bcb3a1; --dim:#80796a; --line:#2a2b3b; --line2:#3d3b4b;
  --accent:#d9b46a; --accent-ink:#1c1406; --accent-soft:rgba(217,180,106,.15);
  --ok:#7fbf98; --ok-soft:rgba(127,191,152,.14); --warn:#e2b55f; --warn-soft:rgba(226,181,95,.14); --bad:#e38b7d; --bad-soft:rgba(227,139,125,.14); --idle:#bcb3a1; --idle-soft:#21263c;
  --shadow:0 16px 40px rgba(0,0,0,.6); color-scheme:dark;
}
:root[data-skin="ouro"] body{background:radial-gradient(48% 38% at 92% -4%,rgba(232,196,119,.20),transparent 70%),radial-gradient(42% 36% at -6% 104%,rgba(120,96,200,.16),transparent 72%),repeating-linear-gradient(45deg,rgba(217,180,106,.028) 0 1px,transparent 1px 22px),repeating-linear-gradient(-45deg,rgba(217,180,106,.028) 0 1px,transparent 1px 22px),var(--bg);background-attachment:fixed}
:root[data-skin="ouro"] .btn.primary,:root[data-skin="ouro"] .ptab[aria-pressed="true"]{background:linear-gradient(135deg,#f7df9e 0%,#d9b46a 42%,#b4862f 100%);border-color:#c99d48;color:#1c1406;box-shadow:0 6px 18px rgba(217,180,106,.22),inset 0 1px 0 rgba(255,255,255,.35)}
:root[data-skin="ouro"] .rn[aria-selected="true"]{background:linear-gradient(90deg,rgba(217,180,106,.18),rgba(217,180,106,.04));color:#f0d48f}
:root[data-skin="ouro"] .rn[aria-selected="true"]::before{background:linear-gradient(180deg,#f7df9e,#b4862f)}
:root[data-skin="ouro"] #tbTitle{background:linear-gradient(90deg,#f7e2a8,#d9b46a 60%,#c29640);-webkit-background-clip:text;background-clip:text;color:transparent}
:root[data-skin="ouro"] .panel,:root[data-skin="ouro"] .kpi,:root[data-skin="ouro"] .mcard{border-color:rgba(217,180,106,.16);box-shadow:inset 0 1px 0 rgba(217,180,106,.10)}
:root[data-skin="ouro"] .topbar,:root[data-skin="ouro"] .rail,:root[data-skin="ouro"] .bottomnav{background:rgba(18,21,36,.92);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px)}
:root[data-skin="claro"]{
  --bg:#f5f5f7; --surface:#ffffff; --sunk:#f0f0f4; --surface3:#e4e4ea; --ink:#1a1a26; --muted:#5a5a6e; --dim:#8b8b9a; --line:#d8d8e0; --line2:#c0c0cc;
  --accent:#7c3aed; --accent-ink:#ffffff; --accent-soft:rgba(124,58,237,.11);
  --ok:#2f8566; --ok-soft:rgba(47,133,102,.14); --warn:#9a6b1f; --warn-soft:rgba(154,107,31,.14); --bad:#b04a4a; --bad-soft:rgba(176,74,74,.13); --idle:#5a5a6e; --idle-soft:#e4e4ea;
  --shadow:0 14px 36px rgba(20,20,40,.18); color-scheme:light;
}
:root[data-skin="oceano"]{
  --bg:#0f141a; --surface:#151b23; --sunk:#1b232d; --surface3:#232d39; --ink:#e2e8f0; --muted:#9aa8b8; --dim:#6b7a8c; --line:#26303c; --line2:#34414f;
  --accent:#6b9bd1; --accent-ink:#0b1320; --accent-soft:rgba(107,155,209,.16);
  --ok:#6fb59a; --ok-soft:rgba(111,181,154,.15); --warn:#d4ad6a; --warn-soft:rgba(212,173,106,.15); --bad:#d98585; --bad-soft:rgba(217,133,133,.15); --idle:#9aa8b8; --idle-soft:#232d39;
  --shadow:0 14px 36px rgba(0,0,0,.45); color-scheme:dark;
}
:root[data-skin="floresta"]{
  --bg:#111412; --surface:#171b18; --sunk:#1d221e; --surface3:#252c27; --ink:#e3e8e4; --muted:#9eaaa2; --dim:#6f7b73; --line:#29302b; --line2:#37403a;
  --accent:#7fb59a; --accent-ink:#0d1a13; --accent-soft:rgba(127,181,154,.16);
  --ok:#8cc0a0; --ok-soft:rgba(140,192,160,.15); --warn:#d2b06b; --warn-soft:rgba(210,176,107,.15); --bad:#d68a82; --bad-soft:rgba(214,138,130,.15); --idle:#9eaaa2; --idle-soft:#252c27;
  --shadow:0 14px 36px rgba(0,0,0,.45); color-scheme:dark;
}
*{box-sizing:border-box}
[hidden]{display:none!important}
body{background:var(--bg);color:var(--ink);font-family:var(--body);font-size:14px;line-height:1.45;-webkit-font-smoothing:antialiased}
body::before{content:"";position:fixed;inset:0 0 auto;height:env(safe-area-inset-top,0px);z-index:30;background:var(--surface)}
.wrap{padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));max-width:1320px;margin:0 auto;width:100%}
.shell{min-height:100%}
.rail{display:none}
.content{min-width:0;display:flex;flex-direction:column}
.topbar{position:sticky;top:env(safe-area-inset-top,0px);z-index:20;background:var(--surface);border-bottom:1px solid var(--line)}
.tb-row{display:flex;align-items:center;gap:10px;min-height:60px;padding-block:8px}
.tb-title{display:flex;align-items:center;gap:12px;margin-right:auto;min-width:0}
.brand-mark{width:40px;height:40px;border-radius:var(--r);background:var(--accent);color:var(--accent-ink);display:grid;place-items:center;font-weight:800;font-size:15px;letter-spacing:-.5px;flex:0 0 auto}
.tb-name{display:flex;flex-direction:column;min-width:0;line-height:1.2}
.tb-name b{font-size:17px;font-weight:800;letter-spacing:-.01em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tb-name span{font-size:12px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.acc-chip{display:inline-block;padding:2px 9px;border-radius:999px;background:var(--accent-soft);color:var(--accent);font-size:11.5px;font-weight:700}
.actions{display:flex;gap:8px;flex-wrap:nowrap}
.iconbtn{width:40px;padding:0;justify-content:center}
.bottomnav{position:fixed;left:0;right:0;bottom:0;z-index:20;display:grid;grid-template-columns:repeat(6,minmax(0,1fr));background:var(--surface);border-top:1px solid var(--line);padding:0 env(safe-area-inset-right,0px) env(safe-area-inset-bottom,0px) env(safe-area-inset-left,0px)}
.bn{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;padding:8px 2px;background:none;border:0;color:var(--muted);font:600 10.5px var(--body);min-height:62px;cursor:pointer;position:relative;min-width:0}
.bn span.t{max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bn svg,.rn svg{width:22px;height:22px;flex:0 0 auto;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.bn[aria-selected="true"]{color:var(--accent)}
.bn .count{position:absolute;top:5px;left:calc(50% + 6px)}
.count{font:700 10.5px var(--body);background:var(--bad);color:#fff;border-radius:999px;padding:1px 6px;min-width:18px;text-align:center}
.rn{display:flex;align-items:center;gap:12px;width:100%;min-height:46px;padding:0 12px;border-radius:var(--r);color:var(--muted);font:600 14px var(--body);background:none;border:0;cursor:pointer;position:relative;text-align:left}
.rn:hover{background:var(--sunk);color:var(--ink)}
.rn[aria-selected="true"]{background:var(--accent-soft);color:var(--accent)}
.rn[aria-selected="true"]::before{content:"";position:absolute;left:-12px;top:9px;bottom:9px;width:4px;border-radius:0 4px 4px 0;background:var(--accent)}
.rn .count{margin-left:auto}
.rail-brand{display:flex;align-items:center;gap:12px;padding:0 6px 22px}
.rail-brand b{display:block;font-size:16px;font-weight:800;letter-spacing:-.01em;line-height:1.15}
.rail-brand span{font-size:11.5px;color:var(--muted)}
.rail-user{margin-top:auto;padding:12px;border-radius:var(--rc);background:var(--sunk);border:1px solid var(--line);display:flex;flex-direction:column;gap:10px}
.rail-user label{font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
main{padding-block:20px calc(90px + env(safe-area-inset-bottom,0px));display:flex;flex-direction:column;gap:18px}
@media (min-width:860px){
  .shell{display:grid;grid-template-columns:236px minmax(0,1fr)}
  .rail{display:flex;flex-direction:column;gap:4px;position:sticky;top:0;height:100vh;height:100dvh;background:var(--surface);border-right:1px solid var(--line);padding:18px 12px calc(18px + env(safe-area-inset-bottom,0px)) max(12px,env(safe-area-inset-left,0px));overflow-y:auto}
  .bottomnav,.tb-title .brand-mark,.actions .mobonly{display:none}
  main{padding-bottom:48px}
}
.btn{font:700 13.5px var(--body);border:1px solid var(--line2);background:var(--surface3);color:var(--ink);border-radius:var(--r);padding:0 16px;min-height:40px;cursor:pointer;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn:hover{border-color:var(--accent)}
.btn.primary{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}
.btn.danger{color:var(--bad);border-color:var(--bad)}
.btn.sm{padding:0 10px;min-height:30px;font-size:12px}
.btn.primary:hover{filter:brightness(1.08)}
.btn:disabled{opacity:.5;cursor:not-allowed}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.banner{border:1px dashed var(--line);background:var(--sunk);border-radius:var(--r);padding:10px 14px;color:var(--muted);font-size:13px}
.banner.err{border-color:var(--bad);color:var(--bad);background:var(--bad-soft)}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.kpi{background:var(--surface);border:1px solid var(--line);border-radius:var(--rc);padding:14px 16px;display:flex;flex-direction:column;gap:4px;position:relative;overflow:hidden}
.kpi::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--c,var(--line))}
.kpi .v{font:600 30px/1.1 var(--mono);font-variant-numeric:tabular-nums}
.kpi .l{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;font-weight:600}
.kpi .s{font-size:12px;color:var(--muted)}
.grid2{display:grid;grid-template-columns:minmax(0,1.3fr) minmax(0,1fr);gap:18px}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:var(--rc);padding:16px;display:flex;flex-direction:column;gap:12px;min-width:0}
.panel h2{font-family:var(--display);font-stretch:85%;font-size:17px;font-weight:700;margin:0;display:flex;justify-content:space-between;align-items:baseline;gap:8px}
.panel h2 small{font:400 12px var(--body);color:var(--muted)}
.toolbar{display:flex;gap:8px;flex-wrap:wrap;align-items:end}
.field{display:flex;flex-direction:column;gap:4px;min-width:0}
.field label{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
input,select,textarea{font:14px var(--body);color:var(--ink);background:var(--sunk);border:1px solid var(--line2);border-radius:var(--r);padding:0 11px;min-height:40px;min-width:0;width:100%}
textarea{padding:9px 11px}
input:focus,select:focus,textarea:focus{border-color:var(--accent);outline:none;box-shadow:0 0 0 3px var(--accent-soft)}
textarea{resize:vertical;min-height:60px}
.toolbar .field{flex:1 1 150px}
.toolbar .field.grow{flex:3 1 220px}
.tablebox{overflow-x:auto;border:1px solid var(--line);border-radius:var(--r);background:var(--surface)}
table{border-collapse:collapse;width:100%;font-size:13px}
th{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);text-align:left;font-weight:600;background:var(--sunk);padding:8px 10px;white-space:nowrap;position:sticky;top:0}
td{padding:8px 10px;border-top:1px solid var(--line);vertical-align:middle}
td.num,th.num{text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums;white-space:nowrap}
td.mono{font-family:var(--mono);font-size:12px;white-space:nowrap}
td.nowrap{white-space:nowrap}
tr.row:hover td{background:var(--sunk)}
.prod{font-weight:600}
td .prod,td.prod{min-width:150px}
.sub{color:var(--muted);font-size:12px}
.pill{display:inline-flex;align-items:center;gap:5px;font:600 11px var(--body);text-transform:uppercase;letter-spacing:.04em;padding:3px 8px;border-radius:999px;white-space:nowrap;background:var(--pbg);color:var(--pfg)}
.pill::before{content:"";width:6px;height:6px;border-radius:50%;background:currentColor}
.st-ok{--pbg:var(--ok-soft);--pfg:var(--ok)} .st-avencer{--pbg:var(--warn-soft);--pfg:var(--warn)} .st-vencida{--pbg:var(--bad-soft);--pfg:var(--bad)}
.st-esgotada,.st-descartada,.st-semval{--pbg:var(--idle-soft);--pfg:var(--idle)}
.mv-entrada{color:var(--ok)} .mv-saida{color:var(--accent)} .mv-descarte{color:var(--bad)} .mv-ajuste{color:var(--warn)} .mv-reclass_out,.mv-reclass_in{color:var(--accent)}
.rowacts{display:flex;gap:4px;justify-content:flex-end;flex-wrap:nowrap}
.empty{padding:32px 16px;text-align:center;color:var(--muted);display:flex;flex-direction:column;gap:10px;align-items:center}
.empty strong{color:var(--ink);font-size:15px}
.chart svg{width:100%;height:auto;display:block}
.chart text{fill:var(--muted);font:11px var(--mono)}
.avg{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}
.avg div{background:var(--sunk);border-radius:6px;padding:8px 10px}
.avg b{display:block;font:600 18px var(--mono);font-variant-numeric:tabular-nums}
.avg span{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
.alist{display:flex;flex-direction:column}
.alist .it{display:grid;grid-template-columns:minmax(0,1fr) auto auto;gap:10px;align-items:center;padding:8px 0;border-top:1px solid var(--line)}
.alist .it:first-child{border-top:0}
.dias{font:600 13px var(--mono);white-space:nowrap}
.lists{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:14px}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{display:inline-flex;align-items:center;gap:4px;background:var(--surface3);color:var(--ink);border-radius:999px;padding:3px 4px 3px 11px;font-size:13px;font-weight:600}
.chip button{border:0;background:none;color:var(--muted);cursor:pointer;font-size:15px;line-height:1;padding:0 5px;border-radius:50%}
.chip button:hover{color:var(--bad)}
.addrow{display:flex;gap:6px}
.overlay{position:fixed;inset:0;background:rgba(10,20,18,.5);z-index:20;display:flex;align-items:flex-start;justify-content:center;overflow-y:auto;padding:max(24px,env(safe-area-inset-top,0px)) max(12px,env(safe-area-inset-right,0px)) 24px max(12px,env(safe-area-inset-left,0px))}
.modal{background:var(--surface);color:var(--ink);border-radius:var(--rc);width:100%;max-width:620px;border:1px solid var(--line);box-shadow:var(--shadow)}
.modal.wide{max-width:900px}
.modal header{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;border-bottom:1px solid var(--line);gap:12px}
.modal header h3{margin:0;font-family:var(--display);font-stretch:85%;font-size:18px}
.modal .body{padding:16px 18px;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
.modal .body .full{grid-column:1/-1}
.modal footer{display:flex;justify-content:flex-end;gap:8px;padding:12px 18px;border-top:1px solid var(--line);flex-wrap:wrap}
.modal .x{border:0;background:none;font-size:22px;color:var(--muted);cursor:pointer;line-height:1}
.ref{background:var(--sunk);border-radius:6px;padding:10px 12px;display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}
.ref .lote{font:600 13px var(--mono)}
.formerr{color:var(--bad);font-size:13px}
.toast{position:fixed;left:50%;transform:translateX(-50%);bottom:calc(20px + env(safe-area-inset-bottom,0px));background:var(--ink);color:var(--bg);padding:10px 16px;border-radius:8px;z-index:30;font-weight:600;font-size:13px;max-width:calc(100% - 32px)}
.muted{color:var(--muted)}
.check{display:flex;gap:6px;align-items:center;font-size:13px;color:var(--muted);white-space:nowrap;padding-bottom:8px}
.check input{width:auto}
@media (max-width:900px){.grid2{grid-template-columns:minmax(0,1fr)}.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:520px){.modal .body{grid-template-columns:minmax(0,1fr)}.avg{grid-template-columns:repeat(2,minmax(0,1fr))}h1{font-size:19px}.actions .btn .lbl{display:none}}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
/* ---------- Evidências (fotos) ---------- */
.fotobox{display:grid;gap:10px;grid-template-columns:1fr}
@media (min-width:600px){.fotobox{grid-template-columns:1fr 1fr}}
.fgrp{background:var(--sunk);border:1px dashed var(--line2);border-radius:12px;padding:10px;display:flex;flex-direction:column;gap:8px;min-width:0}
.fgh{display:flex;justify-content:space-between;align-items:center;gap:8px}
.fthumbs{display:flex;gap:6px;flex-wrap:wrap}
.fthumbs:empty{display:none}
.fth{position:relative;width:64px;height:64px;border-radius:8px;overflow:hidden;background:var(--surface3)}
.fth img{width:100%;height:100%;object-fit:cover;display:block}
.fx{position:absolute;top:2px;right:2px;width:24px;height:24px;border-radius:50%;border:0;background:rgba(0,0,0,.7);color:#fff;cursor:pointer;font-size:15px;line-height:1}
.fgrp label.btn{cursor:pointer;text-transform:none;letter-spacing:0;font:600 14px var(--body);color:var(--ink);flex:1 1 0;justify-content:center}
.evg{display:flex;flex-direction:column;gap:8px;margin-bottom:14px}
.evgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:8px}
.evth{border:1px solid var(--line);background:var(--sunk);border-radius:10px;padding:0;overflow:hidden;cursor:pointer;display:flex;flex-direction:column;color:var(--muted);font:11px var(--body)}
.evth img{width:100%;aspect-ratio:1;object-fit:cover;display:block}
.evth span{padding:4px 6px;text-align:left}
.catb{display:inline-block;font:700 11px var(--body);text-transform:uppercase;letter-spacing:.04em;padding:3px 9px;border-radius:999px;background:color-mix(in srgb,var(--c) 18%,transparent);color:var(--c);white-space:nowrap}
.fstat{font-weight:700;font-size:13px}
.lbx{align-items:center}
.lb{width:min(1100px,100%);display:flex;flex-direction:column;gap:10px}
.lb img{max-width:100%;max-height:78vh;object-fit:contain;border-radius:10px;background:#000;align-self:center}
.lbbar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:10px 12px}
.lbbar .sub{flex:1 1 200px}
@media (min-width:860px){.mlist.cards2{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));align-items:start}}

/* ---------- Pílulas, cartões e ações rápidas (estilo app) ---------- */
.ptabs{display:none;gap:8px;overflow-x:auto;scrollbar-width:none;padding-bottom:2px}
.ptabs::-webkit-scrollbar{display:none}
.ptab{flex:0 0 auto;min-height:44px;padding:0 18px;border-radius:12px;background:var(--sunk);border:1px solid var(--line);color:var(--muted);font:600 15px var(--body);cursor:pointer;white-space:nowrap}
.ptab[aria-pressed="true"]{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}
.mobq{display:none;gap:10px;margin-bottom:14px}
.mlist{display:flex;flex-direction:column;gap:12px}
.mcard{background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--c,var(--accent));border-radius:16px;padding:14px 16px;display:flex;flex-direction:column;gap:8px;min-width:0}
.mrow{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;min-width:0}
.mrow>div{min-width:0}
.mtitle{font-weight:800;font-size:16.5px;line-height:1.25}
.mmeta{color:var(--muted);font-size:13.5px}
.mbig{font:800 22px var(--body);font-variant-numeric:tabular-nums;white-space:nowrap}
.macts{display:flex;gap:8px;flex-wrap:wrap}
.macts .btn{flex:1 1 0;min-width:0}
.ibtn{width:44px;height:44px;border-radius:12px;border:1px solid transparent;background:transparent;color:var(--ink);display:none;place-items:center;cursor:pointer;position:relative;padding:0}
.ibtn svg{width:24px;height:24px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.ibtn:hover{background:var(--sunk)}
.ibtn .count{position:absolute;top:-4px;right:-6px}
.bn .bpill{width:56px;height:32px;border-radius:12px;display:grid;place-items:center;transition:background .15s}
.bn[aria-selected="true"] .bpill{background:var(--accent-soft)}
@media (max-width:859px){
  body{font-size:15px}
  .topbar{border-bottom:0!important;background:linear-gradient(var(--bg) 75%,transparent)!important;-webkit-backdrop-filter:none!important;backdrop-filter:none!important;box-shadow:none!important}
  .tb-row{min-height:68px}
  #tbTitle{font-size:28px;font-weight:800;letter-spacing:-.02em}
  .tb-title .brand-mark,.tb-name span{display:none}
  .actions #btnExport,.actions #btnRetirar,.actions #btnNova,.actions .iconbtn{display:none}
  .actions{gap:4px}
  .ibtn{display:grid}
  .ptabs{display:flex}
  .toolbar .dsk{display:none}
  .mobq{display:flex}
  .mobq .btn{flex:1 1 0;min-height:52px;font-size:16px}
  .btn{min-height:46px;border-radius:12px;font-size:15px}
  .btn.sm{min-height:40px;font-size:14px;border-radius:10px;padding:0 12px}
  input,select{min-height:48px;border-radius:12px;font-size:16px}
  .panel{border-radius:16px}
  .kpi{border-radius:16px;padding:14px}
  .kpi .v{font-size:28px}
  .bn{font-size:11.5px;min-height:68px}
  .bn svg{width:24px;height:24px}
  main{gap:16px}
}

/* Tela de entrada: paisagem ilustrada (visual próprio, igual em qualquer tema) */
#welcome{position:fixed;inset:0;z-index:100;display:grid;place-items:center;overflow:hidden;background:#cfe6f5;transition:opacity .45s ease,visibility .45s}
#welcome.out{opacity:0;visibility:hidden}
#welcome svg.land{position:absolute;inset:0;width:100%;height:100%}
#welcome .wcard{position:relative;width:min(360px,calc(100% - 32px));background:rgba(255,255,255,.86);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.9);border-radius:14px;box-shadow:0 18px 50px rgba(30,50,40,.22);padding:26px 22px 20px;text-align:center;color:#1a1a26;font-family:var(--body);transition:opacity .3s,transform .3s}
#welcome.peek .wcard{opacity:0;transform:translateY(16px);pointer-events:none}
#welcome .wlogo{width:60px;height:60px;border-radius:14px;display:block;margin:0 auto 12px;box-shadow:0 6px 16px rgba(91,33,214,.25)}
#welcome h2{margin:0 0 6px;font-size:20px;font-weight:800;letter-spacing:-.01em;color:#1a1a26}
#welcome .wsub{margin:0 0 14px;color:#4c4c5e;font-size:14px}
#welcome .wrole{display:inline-block;margin-bottom:16px;padding:3px 11px;border-radius:999px;background:rgba(124,58,237,.12);color:#6d28d9;font-size:12px;font-weight:700}
#welcome .wgo{width:100%;min-height:46px;border:0;border-radius:8px;background:#7c3aed;color:#fff;font:700 15px var(--body);cursor:pointer}
#welcome .wgo:hover{background:#6d28d9}
#welcome .wgo:focus-visible{outline:3px solid rgba(124,58,237,.45);outline-offset:2px}
#welcome .wgo{background:linear-gradient(135deg,#f7df9e 0%,#d9b46a 42%,#b4862f 100%);color:#1c1406;box-shadow:0 6px 18px rgba(180,134,47,.3)}
#welcome .wf input:focus{border-color:#c99d48;box-shadow:0 0 0 3px rgba(217,180,106,.25)}
#welcome .wlinks{display:flex;justify-content:center;gap:14px;flex-wrap:wrap}
#welcome .wnote{margin:12px 0 0;font-size:12px;color:#6a6a7e}
#welcome .wf{display:flex;flex-direction:column;gap:5px;text-align:left;margin-bottom:12px}
#welcome .wf label{font-size:12.5px;font-weight:700;color:#3a3a4a}
#welcome .wf input{background:#f4f3f8;border:1px solid #d6d3e2;color:#1a1a26;min-height:46px;border-radius:8px;padding:0 12px;font:15px var(--body);width:100%}
#welcome .wf input:focus{border-color:#7c3aed;box-shadow:0 0 0 3px rgba(124,58,237,.18);outline:none}
#welcome .werr{color:#b03a3a;font-size:13px;min-height:18px;margin:-4px 0 8px;text-align:left}
#welcome .wlink{background:none;border:0;color:#4c4c5e;text-decoration:underline;font:13px var(--body);cursor:pointer;margin-top:12px;padding:4px}
#welcome .wmsg{font-size:14px;color:#3a3a4a;background:#f4f3f8;border-radius:8px;padding:12px;text-align:left;margin:0}
#welcome .wpeek{position:absolute;top:calc(14px + env(safe-area-inset-top,0px));right:calc(14px + env(safe-area-inset-right,0px));display:inline-flex;align-items:center;gap:6px;min-height:38px;padding:0 14px;border-radius:10px;border:1px solid rgba(255,255,255,.9);background:rgba(255,255,255,.86);color:#1a1a26;font:600 13px var(--body);cursor:pointer;box-shadow:0 4px 14px rgba(30,50,40,.12)}
#welcome .cloud{animation:drift 70s linear infinite}
#welcome .cloud.c2{animation-duration:95s;animation-delay:-40s}
@keyframes drift{from{transform:translateX(-260px)}to{transform:translateX(1860px)}}
#welcome .sway{transform-box:fill-box;transform-origin:50% 100%;animation:sway 5s ease-in-out infinite alternate}
@keyframes sway{from{transform:rotate(-1.5deg)}to{transform:rotate(1.5deg)}}
#welcome .blade{transform-box:fill-box;transform-origin:center;animation:spin 9s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
@media (prefers-reduced-motion:reduce){#welcome .cloud,#welcome .sway,#welcome .blade{animation:none}}

</style>
</head>
<body>

<div id="welcome" role="dialog" aria-modal="true" aria-labelledby="wTitle">
  <svg class="land" viewBox="0 0 1600 900" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
<defs><linearGradient id="wsky" x1="0" y1="0" x2="0" y2="1"><stop id="wsky1" offset="0" stop-color="#a9d4f0"/><stop id="wsky2" offset="1" stop-color="#e6f2ec"/></linearGradient>
<radialGradient id="wglow" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#fff3c4" stop-opacity=".9"/><stop offset="1" stop-color="#fff3c4" stop-opacity="0"/></radialGradient></defs>
<rect width="1600" height="900" fill="url(#wsky)"/>
<circle id="wsunglow" cx="1240" cy="150" r="190" fill="url(#wglow)"/>
<circle id="wsun" cx="1240" cy="150" r="88" fill="#ffd97a"/>
<g fill="#ffffff" opacity=".85"><g class="cloud"><ellipse cx="200" cy="140" rx="70" ry="22"/><ellipse cx="240" cy="126" rx="44" ry="24"/></g><g class="cloud c2"><ellipse cx="520" cy="90" rx="58" ry="18"/><ellipse cx="552" cy="78" rx="34" ry="18"/></g></g>
<path d="M0 360 C220 300 420 330 640 300 C880 268 1100 320 1320 290 C1460 272 1540 290 1600 300 V900 H0z" fill="#b9d3cc"/>
<path d="M0 450 C260 400 520 440 760 410 C1000 380 1260 430 1600 395 V900 H0z" fill="#9fc6b4"/>
<g transform="translate(1340 330)"><rect x="-4" y="0" width="8" height="120" fill="#4a5a58"/><g class="blade" ><path d="M0 0 L-6 -46 L6 -46z M0 0 L46 -6 L46 6z M0 0 L6 46 L-6 46z M0 0 L-46 6 L-46 -6z" fill="#f4f4f4" stroke="#4a5a58" stroke-width="2"/></g><circle r="6" fill="#4a5a58"/></g>
<g transform="translate(250 405)"><rect x="0" y="40" width="210" height="80" fill="#eef1ec"/><path d="M-10 44 L105 0 L220 44z" fill="#8a9c96"/><rect x="20" y="66" width="34" height="54" fill="#7c8f88"/><rect x="74" y="64" width="22" height="18" fill="#a9c7d8"/><rect x="108" y="64" width="22" height="18" fill="#a9c7d8"/><rect x="142" y="64" width="22" height="18" fill="#a9c7d8"/><rect x="176" y="64" width="22" height="18" fill="#a9c7d8"/><rect x="160" y="-6" width="18" height="40" fill="#8a9c96"/><rect x="70" y="92" width="120" height="16" rx="3" fill="#7c3aed" opacity=".85"/></g>
<path d="M0 560 C240 520 520 560 760 540 C1040 516 1300 560 1600 520 V900 H0z" fill="#86b79e"/>
<path d="M380 900 C460 760 560 660 700 560" stroke="#e9dcc3" stroke-width="26" fill="none" stroke-linecap="round" opacity=".9"/>
<path d="M700 600 C900 560 1200 590 1600 560 V900 H640z" fill="#c9b48f"/>
<g stroke="#b59f78" stroke-width="3" opacity=".6" fill="none"><path d="M720 660 C950 630 1250 650 1600 620"/><path d="M700 730 C950 700 1260 720 1600 690"/><path d="M690 800 C950 770 1270 790 1600 760"/></g>
<g><ellipse cx="820" cy="643.6" rx="36.0" ry="5.4" fill="#000" opacity=".08"/><path d="M785.8 640 a34.2 27.0 0 0 1 68.4 0z" fill="#2f7a55"/><circle cx="805.6" cy="625.6" r="19.8" fill="#3a8a63"/><circle cx="830.8" cy="622.0" r="21.6" fill="#358259"/><circle cx="800.2" cy="623.8" r="4.5" fill="#d7263d"/><circle cx="814.6" cy="613.0" r="4.5" fill="#d7263d"/><circle cx="832.6" cy="622.0" r="4.5" fill="#d7263d"/><circle cx="841.6" cy="634.6" r="4.5" fill="#d7263d"/><circle cx="807.4" cy="636.4" r="4.5" fill="#d7263d"/><circle cx="823.6" cy="631.0" r="4.5" fill="#d7263d"/><circle cx="794.8" cy="641.8" r="4.5" fill="#d7263d"/><circle cx="834.4" cy="643.6" r="4.5" fill="#d7263d"/></g><g><ellipse cx="940" cy="643.6" rx="36.0" ry="5.4" fill="#000" opacity=".08"/><path d="M905.8 640 a34.2 27.0 0 0 1 68.4 0z" fill="#2f7a55"/><circle cx="925.6" cy="625.6" r="19.8" fill="#3a8a63"/><circle cx="950.8" cy="622.0" r="21.6" fill="#358259"/><circle cx="920.2" cy="623.8" r="4.5" fill="#d7263d"/><circle cx="934.6" cy="613.0" r="4.5" fill="#d7263d"/><circle cx="952.6" cy="622.0" r="4.5" fill="#d7263d"/><circle cx="961.6" cy="634.6" r="4.5" fill="#d7263d"/><circle cx="927.4" cy="636.4" r="4.5" fill="#d7263d"/><circle cx="943.6" cy="631.0" r="4.5" fill="#d7263d"/><circle cx="914.8" cy="641.8" r="4.5" fill="#d7263d"/><circle cx="954.4" cy="643.6" r="4.5" fill="#d7263d"/></g><g><ellipse cx="1060" cy="643.6" rx="36.0" ry="5.4" fill="#000" opacity=".08"/><path d="M1025.8 640 a34.2 27.0 0 0 1 68.4 0z" fill="#2f7a55"/><circle cx="1045.6" cy="625.6" r="19.8" fill="#3a8a63"/><circle cx="1070.8" cy="622.0" r="21.6" fill="#358259"/><circle cx="1040.2" cy="623.8" r="4.5" fill="#d7263d"/><circle cx="1054.6" cy="613.0" r="4.5" fill="#d7263d"/><circle cx="1072.6" cy="622.0" r="4.5" fill="#d7263d"/><circle cx="1081.6" cy="634.6" r="4.5" fill="#d7263d"/><circle cx="1047.4" cy="636.4" r="4.5" fill="#d7263d"/><circle cx="1063.6" cy="631.0" r="4.5" fill="#d7263d"/><circle cx="1034.8" cy="641.8" r="4.5" fill="#d7263d"/><circle cx="1074.4" cy="643.6" r="4.5" fill="#d7263d"/></g><g><ellipse cx="1180" cy="643.6" rx="36.0" ry="5.4" fill="#000" opacity=".08"/><path d="M1145.8 640 a34.2 27.0 0 0 1 68.4 0z" fill="#2f7a55"/><circle cx="1165.6" cy="625.6" r="19.8" fill="#3a8a63"/><circle cx="1190.8" cy="622.0" r="21.6" fill="#358259"/><circle cx="1160.2" cy="623.8" r="4.5" fill="#d7263d"/><circle cx="1174.6" cy="613.0" r="4.5" fill="#d7263d"/><circle cx="1192.6" cy="622.0" r="4.5" fill="#d7263d"/><circle cx="1201.6" cy="634.6" r="4.5" fill="#d7263d"/><circle cx="1167.4" cy="636.4" r="4.5" fill="#d7263d"/><circle cx="1183.6" cy="631.0" r="4.5" fill="#d7263d"/><circle cx="1154.8" cy="641.8" r="4.5" fill="#d7263d"/><circle cx="1194.4" cy="643.6" r="4.5" fill="#d7263d"/></g><g><ellipse cx="1300" cy="643.6" rx="36.0" ry="5.4" fill="#000" opacity=".08"/><path d="M1265.8 640 a34.2 27.0 0 0 1 68.4 0z" fill="#2f7a55"/><circle cx="1285.6" cy="625.6" r="19.8" fill="#3a8a63"/><circle cx="1310.8" cy="622.0" r="21.6" fill="#358259"/><circle cx="1280.2" cy="623.8" r="4.5" fill="#d7263d"/><circle cx="1294.6" cy="613.0" r="4.5" fill="#d7263d"/><circle cx="1312.6" cy="622.0" r="4.5" fill="#d7263d"/><circle cx="1321.6" cy="634.6" r="4.5" fill="#d7263d"/><circle cx="1287.4" cy="636.4" r="4.5" fill="#d7263d"/><circle cx="1303.6" cy="631.0" r="4.5" fill="#d7263d"/><circle cx="1274.8" cy="641.8" r="4.5" fill="#d7263d"/><circle cx="1314.4" cy="643.6" r="4.5" fill="#d7263d"/></g><g><ellipse cx="1420" cy="643.6" rx="36.0" ry="5.4" fill="#000" opacity=".08"/><path d="M1385.8 640 a34.2 27.0 0 0 1 68.4 0z" fill="#2f7a55"/><circle cx="1405.6" cy="625.6" r="19.8" fill="#3a8a63"/><circle cx="1430.8" cy="622.0" r="21.6" fill="#358259"/><circle cx="1400.2" cy="623.8" r="4.5" fill="#d7263d"/><circle cx="1414.6" cy="613.0" r="4.5" fill="#d7263d"/><circle cx="1432.6" cy="622.0" r="4.5" fill="#d7263d"/><circle cx="1441.6" cy="634.6" r="4.5" fill="#d7263d"/><circle cx="1407.4" cy="636.4" r="4.5" fill="#d7263d"/><circle cx="1423.6" cy="631.0" r="4.5" fill="#d7263d"/><circle cx="1394.8" cy="641.8" r="4.5" fill="#d7263d"/><circle cx="1434.4" cy="643.6" r="4.5" fill="#d7263d"/></g><g><ellipse cx="1540" cy="643.6" rx="36.0" ry="5.4" fill="#000" opacity=".08"/><path d="M1505.8 640 a34.2 27.0 0 0 1 68.4 0z" fill="#2f7a55"/><circle cx="1525.6" cy="625.6" r="19.8" fill="#3a8a63"/><circle cx="1550.8" cy="622.0" r="21.6" fill="#358259"/><circle cx="1520.2" cy="623.8" r="4.5" fill="#d7263d"/><circle cx="1534.6" cy="613.0" r="4.5" fill="#d7263d"/><circle cx="1552.6" cy="622.0" r="4.5" fill="#d7263d"/><circle cx="1561.6" cy="634.6" r="4.5" fill="#d7263d"/><circle cx="1527.4" cy="636.4" r="4.5" fill="#d7263d"/><circle cx="1543.6" cy="631.0" r="4.5" fill="#d7263d"/><circle cx="1514.8" cy="641.8" r="4.5" fill="#d7263d"/><circle cx="1554.4" cy="643.6" r="4.5" fill="#d7263d"/></g><g><ellipse cx="880" cy="714.32" rx="43.2" ry="6.48" fill="#000" opacity=".08"/><path d="M838.96 710 a41.040000000000006 32.400000000000006 0 0 1 82.08000000000001 0z" fill="#2f7a55"/><circle cx="862.72" cy="692.72" r="23.76" fill="#3a8a63"/><circle cx="892.96" cy="688.4" r="25.92" fill="#358259"/><circle cx="856.24" cy="690.56" r="5.4" fill="#d7263d"/><circle cx="873.52" cy="677.6" r="5.4" fill="#d7263d"/><circle cx="895.12" cy="688.4" r="5.4" fill="#d7263d"/><circle cx="905.92" cy="703.52" r="5.4" fill="#d7263d"/><circle cx="864.88" cy="705.68" r="5.4" fill="#d7263d"/><circle cx="884.32" cy="699.2" r="5.4" fill="#d7263d"/><circle cx="849.76" cy="712.16" r="5.4" fill="#d7263d"/><circle cx="897.28" cy="714.32" r="5.4" fill="#d7263d"/></g><g><ellipse cx="1000" cy="714.32" rx="43.2" ry="6.48" fill="#000" opacity=".08"/><path d="M958.96 710 a41.040000000000006 32.400000000000006 0 0 1 82.08000000000001 0z" fill="#2f7a55"/><circle cx="982.72" cy="692.72" r="23.76" fill="#3a8a63"/><circle cx="1012.96" cy="688.4" r="25.92" fill="#358259"/><circle cx="976.24" cy="690.56" r="5.4" fill="#d7263d"/><circle cx="993.52" cy="677.6" r="5.4" fill="#d7263d"/><circle cx="1015.12" cy="688.4" r="5.4" fill="#d7263d"/><circle cx="1025.92" cy="703.52" r="5.4" fill="#d7263d"/><circle cx="984.88" cy="705.68" r="5.4" fill="#d7263d"/><circle cx="1004.32" cy="699.2" r="5.4" fill="#d7263d"/><circle cx="969.76" cy="712.16" r="5.4" fill="#d7263d"/><circle cx="1017.28" cy="714.32" r="5.4" fill="#d7263d"/></g><g><ellipse cx="1120" cy="714.32" rx="43.2" ry="6.48" fill="#000" opacity=".08"/><path d="M1078.96 710 a41.040000000000006 32.400000000000006 0 0 1 82.08000000000001 0z" fill="#2f7a55"/><circle cx="1102.72" cy="692.72" r="23.76" fill="#3a8a63"/><circle cx="1132.96" cy="688.4" r="25.92" fill="#358259"/><circle cx="1096.24" cy="690.56" r="5.4" fill="#d7263d"/><circle cx="1113.52" cy="677.6" r="5.4" fill="#d7263d"/><circle cx="1135.12" cy="688.4" r="5.4" fill="#d7263d"/><circle cx="1145.92" cy="703.52" r="5.4" fill="#d7263d"/><circle cx="1104.88" cy="705.68" r="5.4" fill="#d7263d"/><circle cx="1124.32" cy="699.2" r="5.4" fill="#d7263d"/><circle cx="1089.76" cy="712.16" r="5.4" fill="#d7263d"/><circle cx="1137.28" cy="714.32" r="5.4" fill="#d7263d"/></g><g><ellipse cx="1240" cy="714.32" rx="43.2" ry="6.48" fill="#000" opacity=".08"/><path d="M1198.96 710 a41.040000000000006 32.400000000000006 0 0 1 82.08000000000001 0z" fill="#2f7a55"/><circle cx="1222.72" cy="692.72" r="23.76" fill="#3a8a63"/><circle cx="1252.96" cy="688.4" r="25.92" fill="#358259"/><circle cx="1216.24" cy="690.56" r="5.4" fill="#d7263d"/><circle cx="1233.52" cy="677.6" r="5.4" fill="#d7263d"/><circle cx="1255.12" cy="688.4" r="5.4" fill="#d7263d"/><circle cx="1265.92" cy="703.52" r="5.4" fill="#d7263d"/><circle cx="1224.88" cy="705.68" r="5.4" fill="#d7263d"/><circle cx="1244.32" cy="699.2" r="5.4" fill="#d7263d"/><circle cx="1209.76" cy="712.16" r="5.4" fill="#d7263d"/><circle cx="1257.28" cy="714.32" r="5.4" fill="#d7263d"/></g><g><ellipse cx="1360" cy="714.32" rx="43.2" ry="6.48" fill="#000" opacity=".08"/><path d="M1318.96 710 a41.040000000000006 32.400000000000006 0 0 1 82.08000000000001 0z" fill="#2f7a55"/><circle cx="1342.72" cy="692.72" r="23.76" fill="#3a8a63"/><circle cx="1372.96" cy="688.4" r="25.92" fill="#358259"/><circle cx="1336.24" cy="690.56" r="5.4" fill="#d7263d"/><circle cx="1353.52" cy="677.6" r="5.4" fill="#d7263d"/><circle cx="1375.12" cy="688.4" r="5.4" fill="#d7263d"/><circle cx="1385.92" cy="703.52" r="5.4" fill="#d7263d"/><circle cx="1344.88" cy="705.68" r="5.4" fill="#d7263d"/><circle cx="1364.32" cy="699.2" r="5.4" fill="#d7263d"/><circle cx="1329.76" cy="712.16" r="5.4" fill="#d7263d"/><circle cx="1377.28" cy="714.32" r="5.4" fill="#d7263d"/></g><g><ellipse cx="1480" cy="714.32" rx="43.2" ry="6.48" fill="#000" opacity=".08"/><path d="M1438.96 710 a41.040000000000006 32.400000000000006 0 0 1 82.08000000000001 0z" fill="#2f7a55"/><circle cx="1462.72" cy="692.72" r="23.76" fill="#3a8a63"/><circle cx="1492.96" cy="688.4" r="25.92" fill="#358259"/><circle cx="1456.24" cy="690.56" r="5.4" fill="#d7263d"/><circle cx="1473.52" cy="677.6" r="5.4" fill="#d7263d"/><circle cx="1495.12" cy="688.4" r="5.4" fill="#d7263d"/><circle cx="1505.92" cy="703.52" r="5.4" fill="#d7263d"/><circle cx="1464.88" cy="705.68" r="5.4" fill="#d7263d"/><circle cx="1484.32" cy="699.2" r="5.4" fill="#d7263d"/><circle cx="1449.76" cy="712.16" r="5.4" fill="#d7263d"/><circle cx="1497.28" cy="714.32" r="5.4" fill="#d7263d"/></g><g><ellipse cx="1600" cy="714.32" rx="43.2" ry="6.48" fill="#000" opacity=".08"/><path d="M1558.96 710 a41.040000000000006 32.400000000000006 0 0 1 82.08000000000001 0z" fill="#2f7a55"/><circle cx="1582.72" cy="692.72" r="23.76" fill="#3a8a63"/><circle cx="1612.96" cy="688.4" r="25.92" fill="#358259"/><circle cx="1576.24" cy="690.56" r="5.4" fill="#d7263d"/><circle cx="1593.52" cy="677.6" r="5.4" fill="#d7263d"/><circle cx="1615.12" cy="688.4" r="5.4" fill="#d7263d"/><circle cx="1625.92" cy="703.52" r="5.4" fill="#d7263d"/><circle cx="1584.88" cy="705.68" r="5.4" fill="#d7263d"/><circle cx="1604.32" cy="699.2" r="5.4" fill="#d7263d"/><circle cx="1569.76" cy="712.16" r="5.4" fill="#d7263d"/><circle cx="1617.28" cy="714.32" r="5.4" fill="#d7263d"/></g><g><ellipse cx="820" cy="785.04" rx="50.4" ry="7.5600000000000005" fill="#000" opacity=".08"/><path d="M772.12 780 a47.88 37.8 0 0 1 95.76 0z" fill="#2f7a55"/><circle cx="799.84" cy="759.84" r="27.72" fill="#3a8a63"/><circle cx="835.12" cy="754.8" r="30.240000000000002" fill="#358259"/><circle cx="792.28" cy="757.32" r="6.3" fill="#d7263d"/><circle cx="812.44" cy="742.2" r="6.3" fill="#d7263d"/><circle cx="837.64" cy="754.8" r="6.3" fill="#d7263d"/><circle cx="850.24" cy="772.44" r="6.3" fill="#d7263d"/><circle cx="802.36" cy="774.96" r="6.3" fill="#d7263d"/><circle cx="825.04" cy="767.4" r="6.3" fill="#d7263d"/><circle cx="784.72" cy="782.52" r="6.3" fill="#d7263d"/><circle cx="840.16" cy="785.04" r="6.3" fill="#d7263d"/></g><g><ellipse cx="940" cy="785.04" rx="50.4" ry="7.5600000000000005" fill="#000" opacity=".08"/><path d="M892.12 780 a47.88 37.8 0 0 1 95.76 0z" fill="#2f7a55"/><circle cx="919.84" cy="759.84" r="27.72" fill="#3a8a63"/><circle cx="955.12" cy="754.8" r="30.240000000000002" fill="#358259"/><circle cx="912.28" cy="757.32" r="6.3" fill="#d7263d"/><circle cx="932.44" cy="742.2" r="6.3" fill="#d7263d"/><circle cx="957.64" cy="754.8" r="6.3" fill="#d7263d"/><circle cx="970.24" cy="772.44" r="6.3" fill="#d7263d"/><circle cx="922.36" cy="774.96" r="6.3" fill="#d7263d"/><circle cx="945.04" cy="767.4" r="6.3" fill="#d7263d"/><circle cx="904.72" cy="782.52" r="6.3" fill="#d7263d"/><circle cx="960.16" cy="785.04" r="6.3" fill="#d7263d"/></g><g><ellipse cx="1060" cy="785.04" rx="50.4" ry="7.5600000000000005" fill="#000" opacity=".08"/><path d="M1012.12 780 a47.88 37.8 0 0 1 95.76 0z" fill="#2f7a55"/><circle cx="1039.84" cy="759.84" r="27.72" fill="#3a8a63"/><circle cx="1075.12" cy="754.8" r="30.240000000000002" fill="#358259"/><circle cx="1032.28" cy="757.32" r="6.3" fill="#d7263d"/><circle cx="1052.44" cy="742.2" r="6.3" fill="#d7263d"/><circle cx="1077.64" cy="754.8" r="6.3" fill="#d7263d"/><circle cx="1090.24" cy="772.44" r="6.3" fill="#d7263d"/><circle cx="1042.36" cy="774.96" r="6.3" fill="#d7263d"/><circle cx="1065.04" cy="767.4" r="6.3" fill="#d7263d"/><circle cx="1024.72" cy="782.52" r="6.3" fill="#d7263d"/><circle cx="1080.16" cy="785.04" r="6.3" fill="#d7263d"/></g><g><ellipse cx="1180" cy="785.04" rx="50.4" ry="7.5600000000000005" fill="#000" opacity=".08"/><path d="M1132.12 780 a47.88 37.8 0 0 1 95.76 0z" fill="#2f7a55"/><circle cx="1159.84" cy="759.84" r="27.72" fill="#3a8a63"/><circle cx="1195.12" cy="754.8" r="30.240000000000002" fill="#358259"/><circle cx="1152.28" cy="757.32" r="6.3" fill="#d7263d"/><circle cx="1172.44" cy="742.2" r="6.3" fill="#d7263d"/><circle cx="1197.64" cy="754.8" r="6.3" fill="#d7263d"/><circle cx="1210.24" cy="772.44" r="6.3" fill="#d7263d"/><circle cx="1162.36" cy="774.96" r="6.3" fill="#d7263d"/><circle cx="1185.04" cy="767.4" r="6.3" fill="#d7263d"/><circle cx="1144.72" cy="782.52" r="6.3" fill="#d7263d"/><circle cx="1200.16" cy="785.04" r="6.3" fill="#d7263d"/></g><g><ellipse cx="1300" cy="785.04" rx="50.4" ry="7.5600000000000005" fill="#000" opacity=".08"/><path d="M1252.12 780 a47.88 37.8 0 0 1 95.76 0z" fill="#2f7a55"/><circle cx="1279.84" cy="759.84" r="27.72" fill="#3a8a63"/><circle cx="1315.12" cy="754.8" r="30.240000000000002" fill="#358259"/><circle cx="1272.28" cy="757.32" r="6.3" fill="#d7263d"/><circle cx="1292.44" cy="742.2" r="6.3" fill="#d7263d"/><circle cx="1317.64" cy="754.8" r="6.3" fill="#d7263d"/><circle cx="1330.24" cy="772.44" r="6.3" fill="#d7263d"/><circle cx="1282.36" cy="774.96" r="6.3" fill="#d7263d"/><circle cx="1305.04" cy="767.4" r="6.3" fill="#d7263d"/><circle cx="1264.72" cy="782.52" r="6.3" fill="#d7263d"/><circle cx="1320.16" cy="785.04" r="6.3" fill="#d7263d"/></g><g><ellipse cx="1420" cy="785.04" rx="50.4" ry="7.5600000000000005" fill="#000" opacity=".08"/><path d="M1372.12 780 a47.88 37.8 0 0 1 95.76 0z" fill="#2f7a55"/><circle cx="1399.84" cy="759.84" r="27.72" fill="#3a8a63"/><circle cx="1435.12" cy="754.8" r="30.240000000000002" fill="#358259"/><circle cx="1392.28" cy="757.32" r="6.3" fill="#d7263d"/><circle cx="1412.44" cy="742.2" r="6.3" fill="#d7263d"/><circle cx="1437.64" cy="754.8" r="6.3" fill="#d7263d"/><circle cx="1450.24" cy="772.44" r="6.3" fill="#d7263d"/><circle cx="1402.36" cy="774.96" r="6.3" fill="#d7263d"/><circle cx="1425.04" cy="767.4" r="6.3" fill="#d7263d"/><circle cx="1384.72" cy="782.52" r="6.3" fill="#d7263d"/><circle cx="1440.16" cy="785.04" r="6.3" fill="#d7263d"/></g><g><ellipse cx="1540" cy="785.04" rx="50.4" ry="7.5600000000000005" fill="#000" opacity=".08"/><path d="M1492.12 780 a47.88 37.8 0 0 1 95.76 0z" fill="#2f7a55"/><circle cx="1519.84" cy="759.84" r="27.72" fill="#3a8a63"/><circle cx="1555.12" cy="754.8" r="30.240000000000002" fill="#358259"/><circle cx="1512.28" cy="757.32" r="6.3" fill="#d7263d"/><circle cx="1532.44" cy="742.2" r="6.3" fill="#d7263d"/><circle cx="1557.64" cy="754.8" r="6.3" fill="#d7263d"/><circle cx="1570.24" cy="772.44" r="6.3" fill="#d7263d"/><circle cx="1522.36" cy="774.96" r="6.3" fill="#d7263d"/><circle cx="1545.04" cy="767.4" r="6.3" fill="#d7263d"/><circle cx="1504.72" cy="782.52" r="6.3" fill="#d7263d"/><circle cx="1560.16" cy="785.04" r="6.3" fill="#d7263d"/></g>
<g class="sway"><rect x="118" y="560" width="10" height="120" fill="#5b4a3a"/><circle cx="123" cy="540" r="52" fill="#2f6f4f"/><circle cx="90" cy="570" r="38" fill="#2a6447"/><circle cx="158" cy="572" r="40" fill="#2a6447"/></g>
<g class="sway" style="animation-delay:-2s"><rect x="560" y="520" width="8" height="80" fill="#5b4a3a"/><circle cx="564" cy="506" r="34" fill="#3a7d5c"/></g>
<path d="M0 760 C200 720 420 760 620 800 C700 818 720 860 700 900 H0z" fill="#6aa68a"/>
<g transform="translate(330 735)"><rect x="-70" y="10" width="120" height="44" rx="8" fill="#ffffff"/><rect x="50" y="22" width="40" height="32" rx="6" fill="#7c3aed"/><rect x="58" y="28" width="18" height="12" rx="2" fill="#cfe2f3"/><circle cx="-40" cy="58" r="12" fill="#2b2b36"/><circle cx="62" cy="58" r="12" fill="#2b2b36"/><rect x="-58" y="20" width="96" height="12" rx="3" fill="#e9e2ff"/></g>
<g fill="#4f8a6c" opacity=".8"><path d="M40 860 q6 -30 10 0z M56 860 q8 -36 12 0z M74 860 q4 -24 8 0z"/><path d="M1500 860 q6 -30 10 0z M1516 860 q8 -36 12 0z"/></g>
<rect id="wnight" x="0" y="0" width="1600" height="900" fill="#0b1630" opacity="0"/>
</svg>
  <button type="button" class="wpeek" id="wPeek" aria-pressed="false"><svg width="16" height="16" viewBox="0 0 24 24" aria-hidden="true" style="fill:none;stroke:currentColor;stroke-width:2"><path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z"/><circle cx="12" cy="12" r="3"/></svg><span>Ver a paisagem</span></button>
  <div class="wcard">
    <img class="wlogo" src="data:image/png;base64,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" alt="">
    <h2 id="wTitle">Olá!</h2>
    <p class="wsub">Controle de amostras de produto acabado</p>
    <div><span class="wrole" id="wRole" hidden></span></div>
    <form id="wForm" novalidate autocomplete="on"><div id="wBody"><p class="wnote">Verificando seu acesso…</p></div></form>
  </div>
</div>
<div class="shell">
  <aside class="rail" aria-label="Menu">
    <div class="rail-brand"><div class="brand-mark" aria-hidden="true">AP</div><div><b>Amostras PA</b><span>Controle de Qualidade</span></div></div>
    <nav id="tabs" role="tablist" style="display:flex;flex-direction:column;gap:4px">
      <button class="rn nv" role="tab" data-tab="painel" aria-selected="false"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/></svg><span>Painel</span></button>
      <button class="rn nv" role="tab" data-tab="estoque" aria-selected="false"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg><span>Estoque</span></button>
      <button class="rn nv" role="tab" data-tab="movs" aria-selected="false"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 4v16"/><path d="M3 8l4-4 4 4"/><path d="M17 20V4"/><path d="M21 16l-4 4-4-4"/></svg><span>Movimentações</span></button>
      <button class="rn nv" role="tab" data-tab="evidencias" aria-selected="false"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg><span>Evidências</span></button>
      <button class="rn nv" role="tab" data-tab="vencidas" aria-selected="false"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2"/><path d="M5 3L2 6"/><path d="M19 3l3 3"/></svg><span>Vencidas</span><span class="count cntVenc" hidden>0</span></button>
      <button class="rn nv" role="tab" data-tab="clientes" aria-selected="false"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.8-3.6 3.4-5.5 6.5-5.5s5.7 1.9 6.5 5.5"/><path d="M16 4.6a3.5 3.5 0 0 1 0 6.8"/><path d="M18.5 14.8c1.6.8 2.6 2.4 3 5.2"/></svg><span>Clientes / Lab. externo</span></button>
      <button class="rn nv" role="tab" data-tab="cadastros" aria-selected="false"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h10"/><path d="M18 6h2"/><circle cx="16" cy="6" r="2"/><path d="M4 12h3"/><path d="M11 12h9"/><circle cx="9" cy="12" r="2"/><path d="M4 18h12"/><path d="M20 18h0"/><circle cx="18" cy="18" r="2"/></svg><span>Cadastros</span></button>
    </nav>
    <div class="rail-user">
      <div><span class="acc-chip" id="accChip" hidden></span></div>
      <button type="button" class="btn" id="btnSair" hidden>Sair</button>
      <div class="field"><label for="skinSel">Tema</label><select id="skinSel"><option value="ouro">Dourado</option><option value="roxo">Roxo escuro</option><option value="claro">Claro</option><option value="oceano">Oceano</option><option value="floresta">Floresta</option></select></div>
    </div>
  </aside>
  <div class="content">
    <header class="topbar">
      <div class="wrap tb-row">
        <div class="tb-title"><div class="brand-mark" aria-hidden="true">AP</div><div class="tb-name"><b id="tbTitle">Painel</b><span id="tbSub">Amostras de Produto Acabado</span></div></div>
        <div class="actions">
          <button class="btn iconbtn mobonly" id="btnSkin" title="Trocar tema" aria-label="Trocar tema" style="font-size:18px">◐</button>
          <button class="ibtn" id="ibExport" title="Exportar Excel" aria-label="Exportar Excel"><svg viewBox="0 0 24 24"><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M4 21h16"/></svg></button>
          <button class="ibtn" id="ibBell" title="Vencidas" aria-label="Ver vencidas"><svg viewBox="0 0 24 24"><path d="M6 8a6 6 0 1 1 12 0c0 7 3 8 3 8H3s3-1 3-8"/><path d="M10 20a2 2 0 0 0 4 0"/></svg><span class="count cntVenc" hidden>0</span></button>
          <button class="ibtn" id="ibSkin" title="Trocar tema" aria-label="Trocar tema"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 3a9 9 0 0 0 0 18z" style="fill:currentColor"/></svg></button>
          <button class="ibtn" id="ibSair" title="Sair" aria-label="Sair" hidden><svg viewBox="0 0 24 24"><path d="M12 3v9"/><path d="M6.4 6.6a8 8 0 1 0 11.2 0"/></svg></button>
          <button class="btn iconbtn mobonly" id="btnSair2" title="Sair" aria-label="Sair" hidden style="font-size:16px">⏻</button>
          <button class="btn" id="btnExport" title="Exportar para Excel">⭳ <span class="lbl">Exportar Excel</span></button>
          <button class="btn" id="btnRetirar" title="Registrar retirada">− <span class="lbl">Registrar retirada</span></button>
          <button class="btn primary" id="btnNova" title="Nova amostra">+ <span class="lbl">Nova amostra</span></button>
        </div>
      </div>
    </header>
    <main class="wrap" id="main"></main>
  </div>
</div>
<nav class="bottomnav" id="bnav" role="tablist" aria-label="Navegação">
  <button class="bn nv" role="tab" data-tab="painel" aria-selected="false" aria-label="Painel"><span class="bpill"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/></svg></span><span class="t">Painel</span></button>
  <button class="bn nv" role="tab" data-tab="estoque" aria-selected="false" aria-label="Estoque"><span class="bpill"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg></span><span class="t">Estoque</span></button>
  <button class="bn nv" role="tab" data-tab="movs" aria-selected="false" aria-label="Movimentações"><span class="bpill"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 4v16"/><path d="M3 8l4-4 4 4"/><path d="M17 20V4"/><path d="M21 16l-4 4-4-4"/></svg></span><span class="t">Movim.</span></button>
  <button class="bn nv" role="tab" data-tab="evidencias" aria-selected="false" aria-label="Evidências"><span class="bpill"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg></span><span class="t">Evidências</span></button>
  <button class="bn nv" role="tab" data-tab="clientes" aria-selected="false" aria-label="Clientes / Lab. externo"><span class="bpill"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.8-3.6 3.4-5.5 6.5-5.5s5.7 1.9 6.5 5.5"/><path d="M16 4.6a3.5 3.5 0 0 1 0 6.8"/><path d="M18.5 14.8c1.6.8 2.6 2.4 3 5.2"/></svg></span><span class="t">Clientes</span></button>
  <button class="bn nv" role="tab" data-tab="cadastros" aria-selected="false" aria-label="Cadastros"><span class="bpill"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h10"/><path d="M18 6h2"/><circle cx="16" cy="6" r="2"/><path d="M4 12h3"/><path d="M11 12h9"/><circle cx="9" cy="12" r="2"/><path d="M4 18h12"/><path d="M20 18h0"/><circle cx="18" cy="18" r="2"/></svg></span><span class="t">Cadastros</span></button>
</nav>

<script>
/* ======================= Estado ======================= */
const DEFAULT_CFG = {
  alerta: 60,
  analistas: [],
  locais: [],
  produtos: [],
  destinos: ["Indústria","Agrícola","Externo"],
  tipos: ["Retenção (contraprova)","Análise","Estabilidade","Cliente / comercial","Reclamação"],
  motivos: ["Análise de rotina","Análise de reclamação","Análise externa (laboratório)","Envio a cliente","Estudo de estabilidade","Auditoria / fiscalização","Descarte – vencida","Descarte – avariada","Outro"],
  unidades: ["un","fr","cx","pct","kg","g","L","mL"]
};
const S = { role:"total", isOwner:false, myId:null, canW:null, acessos:{}, pessoas:[], amostras:[], movs:[], hist:[], clientes:[], envh:[], envIds:{}, stkIds:{}, stkKg:{}, envKg:{}, fc:{q:"",pais:"",tipo:""}, cfg:{...DEFAULT_CFG}, mode:"loading", tab:"painel", readonly:false,
  f:{ q:"", status:"ativas", local:"", tipo:"", sort:"validade" },
  fm:{ q:"", tipo:"", resp:"", de:"", ate:"" },
  fe:{ cat:"", q:"", foto:"", de:"", ate:"" } };
let db=null, me=null, pendingRefresh=false;
const FULL=()=>S.role==="total";
const ROLE_LABEL={total:"Acesso total",operador:"Inclusão e saída",leitura:"Visualização",bloqueado:"Sem acesso",pendente:"Aguardando liberação"};
function computeRole(){
  if(S.mode!=="db"){ S.role="total"; S.readonly=false; return; }
  if(S.isOwner) S.role="total";
  else { const r=S.acessos[S.myId]; S.role=(r==="total"||r==="operador"||r==="bloqueado"||r==="leitura")?r:"pendente"; }
  S.readonly = S.role!=="total" && S.role!=="operador";
}

const LS = { get(k,d){try{const v=localStorage.getItem("apa:"+k);return v==null?d:JSON.parse(v)}catch(e){return d}},
  set(k,v){try{localStorage.setItem("apa:"+k,JSON.stringify(v))}catch(e){}} };
S.tab = LS.get("tab","painel");

/* ======================= Utilidades ======================= */
const $ = s=>document.querySelector(s);
const esc = v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pad = n=>String(n).padStart(2,"0");
const isoLocal = d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const hoje = ()=>isoLocal(new Date());
const parseISO = s=>{ if(!s) return null; const [y,m,d]=s.split("-").map(Number); return new Date(y,m-1,d); };
const fmtD = s=>{ if(!s) return "—"; const [y,m,d]=s.split("-"); return `${d}/${m}/${y}`; };
const dias = s=>{ const d=parseISO(s); if(!d) return null; const t=parseISO(hoje()); return Math.round((d-t)/864e5); };
const nf = n=>(Math.round(n*100)/100).toLocaleString("pt-BR");
const uid = ()=>Date.now().toString(36)+Math.random().toString(36).slice(2,8);
const norm = s=>String(s??"").normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().trim();
const PRODS = ()=>(S.cfg.produtos||[]).map(p=>typeof p==="string"?{nome:p,mapa:"",categoria:""}:p);
const mapaDe = nome=>{ const p=PRODS().find(x=>norm(x.nome)===norm(nome)); return p?p.mapa||"":""; };
const MES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];

function toast(msg){ const t=document.createElement("div"); t.className="toast"; t.textContent=msg; document.body.appendChild(t); setTimeout(()=>t.remove(),2600); }

/* ======================= Cálculos ======================= */
function saldos(){
  const m={};
  for(const v of S.movs){ const q=Number(v.qtd)||0; const s=(v.tipo==="entrada"||v.tipo==="ajuste"||v.tipo==="reclass_in")?q:-q; m[v.amostraId]=(m[v.amostraId]||0)+s; }
  return m;
}
function ultimaMov(){ const m={}; for(const v of S.movs){ if(!m[v.amostraId]||v.data>m[v.amostraId].data) m[v.amostraId]=v; } return m; }
function statusOf(a,saldo){
  if(a.descartada) return "descartada";
  if(saldo<=0) return "esgotada";
  if(!a.validade) return "semval";
  const d=dias(a.validade);
  if(d<=0) return "vencida";
  if(d<=Number(S.cfg.alerta||60)) return "avencer";
  return "ok";
}
const ST_LABEL = {ok:"Ok",avencer:"A vencer",vencida:"Vencida",esgotada:"Esgotada",descartada:"Descartada",semval:"Sem validade"};
const ST_ORDER = {vencida:0,avencer:1,semval:2,ok:3,esgotada:4,descartada:5};
const pill = st=>`<span class="pill st-${st}">${ST_LABEL[st]}</span>`;
function rows(){
  const sd=saldos(), um=ultimaMov();
  return S.amostras.map(a=>{ const saldo=sd[a.id]||0; const st=statusOf(a,saldo); const u=um[a.id];
    const atual=[a.atualizadoEm,u&&u.data].filter(Boolean).sort().pop()||"";
    return {...a,saldo,st,dias:dias(a.validade),atual,respUlt:(u&&u.responsavel)||a.responsavel||""}; });
}
/* Retiradas usadas nas médias: lançamentos do app + histórico mensal importado da planilha */
function saidasAll(){
  const h=S.hist.map(x=>({produto:x.produto,data:(x.mes||"")+"-01",qtd:Number(x.qtd)||0,hist:true}));
  return [...S.movs.filter(m=>m.tipo==="saida"), ...h];
}
/* IDs (posições) em estoque de cada lote: campo do próprio lote, ou o levantamento importado da planilha */
function idsOf(a){
  if(!a) return {ids:[],q:{},conf:"",kg:{}};
  if(Array.isArray(a.ids)) return {ids:a.ids,q:a.idsQ||{},conf:a.idsConf||"",kg:a.idsKg||{}};
  const m=S.stkIds[a.id]; return m?{ids:m.ids||[],q:m.q||{},conf:m.conf||"",kg:S.stkKg[a.id]||{}}:{ids:[],q:{},conf:"",kg:{}};
}
const kgOf = a=>{ const io=idsOf(a); return io.ids.reduce((s,i)=>s+(Number(io.kg[i])||0),0); };
const fkg = v=>(Math.round(v*1000)/1000).toLocaleString("pt-BR",{minimumFractionDigits:v>=100?1:3,maximumFractionDigits:3});
const parseIds = s=>[...new Set(String(s||"").split(/[,;\s]+/).map(x=>x.trim().toUpperCase()).filter(Boolean))];
function parseKg(t){
  const toks=String(t||"").replace(/kg/gi," ").split(/[\n\r\t;]+|\s{2,}|\s(?=\d)/).map(x=>x.trim()).filter(Boolean);
  if(!toks.length) return null; let sum=0;
  for(let x of toks){ x=x.replace(/\s/g,""); if(x.includes(",")&&x.includes(".")) x=x.replace(/\./g,"").replace(",","."); else x=x.replace(",",".");
    if(!/^-?\d*\.?\d+$/.test(x)) return null; const v=Number(x); if(!(v>=0)) return null; sum+=v; }
  return Math.round(sum*1000)/1000;
}
const idLabel = (i,q,kg)=>(q&&q[i]>1?`${i}×${q[i]}`:i)+(kg&&kg[i]?` · ${fkg(kg[i])} kg`:"");
const MV_LABEL = {entrada:"Entrada",saida:"Retirada",descarte:"Descarte",ajuste:"Ajuste",reclass_out:"Reclassificação (saída)",reclass_in:"Reclassificação (entrada)"};

/* ======================= Persistência ======================= */
const store = {
  async addAmostra(a, qtd, resp, data){
    const id=uid(); const doc={...a, criadoEm:new Date().toISOString(), atualizadoEm:data||hoje(), descartada:false};
    if(db){ await db.collection("amostras").doc(id).set(doc); }
    else { S.amostras.push({id,...doc}); }
    if(qtd>0) await store.addMov({amostraId:id, produto:a.produto, lote:a.lote, tipo:"entrada", qtd, data:data||hoje(), motivo:"Cadastro inicial", destino:"", responsavel:resp||a.responsavel||"", obs:""}, true);
    else if(!db) refresh();
    return id;
  },
  async updateAmostra(id, patch){
    if(db){ await db.collection("amostras").doc(id).update(patch); }
    else { const a=S.amostras.find(x=>x.id===id); Object.assign(a,patch); refresh(); }
  },
  async deleteAmostra(id){
    const mine=S.movs.filter(m=>m.amostraId===id);
    if(db){ for(const m of mine) await db.collection("movs").doc(m.id).delete(); await db.collection("amostras").doc(id).delete(); }
    else { S.movs=S.movs.filter(m=>m.amostraId!==id); S.amostras=S.amostras.filter(a=>a.id!==id); refresh(); }
  },
  async addMov(m, quiet){
    const id=uid(); const doc={...m, criadoEm:new Date().toISOString()};
    if(db){ await db.collection("movs").doc(id).set(doc); }
    else { S.movs.push({id,...doc}); if(!quiet) refresh(); else refresh(); }
    return id;
  },
  async deleteMov(id){
    if(db){ try{ const q=await db.collection("evidencias").where("movId","==",id).get(); for(const d of q.docs) await db.collection("evidencias").doc(d.id).delete(); }catch(e){ console.error(e); } }
    if(db) await db.collection("movs").doc(id).delete();
    else { S.movs=S.movs.filter(m=>m.id!==id); refresh(); }
  },
  async saveCfg(cfg){
    S.cfg=cfg;
    if(db){ await db.doc("config/listas").set(cfg); } else refresh();
  }
};
async function guard(fn, okMsg){
  try{ await fn(); if(okMsg) toast(okMsg); return true; }
  catch(e){
    const code=e&&e.code;
    if((code==="invalid_argument"||code==="permission-denied") && S.mode==="db"){ toast("Seu nível de acesso não permite essa alteração."); }
    else if(code==="quota_exceeded"||code==="resource-exhausted") toast("O banco de dados atingiu o limite. Exporte e exclua registros antigos.");
    else toast("Não foi possível salvar agora. Tente novamente.");
    console.error(e); return false;
  }
}

/* ======================= Inicialização ======================= */
async function init(){
  render();
  const cfgOk=window.firebase&&window.FIREBASE_CONFIG&&FIREBASE_CONFIG.apiKey&&!/COLE/i.test(FIREBASE_CONFIG.apiKey);
  if(!cfgOk){ S.mode="local"; render(); WEL.fill(); W_MODE="demo"; wBody(`<p class="wmsg" style="margin-bottom:12px">O Firebase ainda não foi configurado. Preencha o arquivo <b>config.js</b> com os dados do seu projeto (veja o passo a passo).</p><button type="submit" class="wgo">Ver em modo demonstração</button>`); return; }
  firebase.initializeApp(FIREBASE_CONFIG);
  db=firebase.firestore();
  try{ await db.enablePersistence({synchronizeTabs:true}); }catch(e){}
  firebase.auth().onAuthStateChanged(u=>onUser(u));
}

/* ======================= Render ======================= */
function setTab(t){ S.tab=t; LS.set("tab",t); render(); }
["tabs","bnav"].forEach(n=>document.getElementById(n).addEventListener("click",e=>{ const b=e.target.closest("button[data-tab]"); if(b){ setTab(b.dataset.tab); window.scrollTo(0,0); } }));
const TAB_LABEL={painel:"Painel",estoque:"Estoque",movs:"Movimentações",evidencias:"Evidências",vencidas:"Vencidas",clientes:"Clientes / Lab. externo",cadastros:"Cadastros"};
const SKINS=["ouro","roxo","claro","oceano","floresta"];
function applySkin(k){ if(!SKINS.includes(k)) k="ouro"; if(k==="roxo") delete document.documentElement.dataset.skin; else document.documentElement.dataset.skin=k; const mt=document.querySelector('meta[name="theme-color"]'); if(mt) mt.content=getComputedStyle(document.documentElement).getPropertyValue("--surface").trim()||"#121524"; const sel=document.getElementById("skinSel"); if(sel) sel.value=k; LS.set("skin",k); }
applySkin(LS.get("skin","ouro"));
document.getElementById("skinSel").addEventListener("change",e=>applySkin(e.target.value));
document.getElementById("ibSkin").addEventListener("click",()=>document.getElementById("btnSkin").click());
document.getElementById("ibBell").addEventListener("click",()=>{ setTab("vencidas"); window.scrollTo(0,0); });
document.getElementById("ibExport").addEventListener("click",()=>exportar());
document.getElementById("ibSair").addEventListener("click",()=>sairApp());
window.matchMedia("(max-width:859px)").addEventListener("change",()=>render());
document.getElementById("btnSkin").addEventListener("click",()=>{ const cur=LS.get("skin","ouro"); applySkin(SKINS[(SKINS.indexOf(cur)+1)%SKINS.length]); toast("Tema: "+document.querySelector(`#skinSel option[value="${LS.get("skin","ouro")}"]`).textContent); });

function render(){
  document.querySelectorAll("button.nv[data-tab]").forEach(b=>b.setAttribute("aria-selected", String(b.dataset.tab===S.tab)));
  const tt=document.getElementById("tbTitle"); if(tt) tt.textContent=(isMobile()?{movs:"Movimentos",clientes:"Clientes"}[S.tab]:"")||TAB_LABEL[S.tab]||"";
  const ts=document.getElementById("tbSub"); if(ts) ts.textContent="Amostras de Produto Acabado"+(S.mode==="db"?" · "+ROLE_LABEL[S.role]:"");
  const banner = S.mode==="local" ? `<div class="banner">Modo demonstração: este navegador não está conectado ao banco compartilhado, então nada do que for lançado aqui será salvo. Preencha o arquivo config.js com os dados do Firebase para usar com a equipe.</div>`
    : S.mode==="error" ? `<div class="banner err">A conexão com o banco foi interrompida. Recarregue a página para continuar.</div>`
    : S.mode==="auth" ? ""
    : S.role==="leitura" ? `<div class="banner">Seu acesso é de visualização. Para lançar movimentações, peça acesso à proprietária do controle.</div>`
    : S.canW===false ? `<div class="banner">Seu acesso no app é “${ROLE_LABEL[S.role]}”, mas o link foi compartilhado com você só para visualização. Peça à proprietária para compartilhar como Editor.</div>` : "";
  const shells = {evidencias:shellEvid, painel:shellPainel, estoque:shellEstoque, movs:shellMovs, vencidas:shellVencidas, clientes:shellClientes, cadastros:shellCadastros};
  $("#main").innerHTML = banner + shells[S.tab]();
  bindShell();
  refresh(true);
}
function refresh(force){
  const ae=document.activeElement;
  if(!force && ae && ae.closest && ae.closest("#data") && /INPUT|SELECT|TEXTAREA/.test(ae.tagName)){ pendingRefresh=true; return; }
  pendingRefresh=false;
  const vc = rows().filter(r=>r.st==="vencida").length;
  document.querySelectorAll(".cntVenc").forEach(c=>{ c.hidden=!vc; c.textContent=vc; });
  const el=$("#data"); if(!el) return;
  const datas={evidencias:dataEvid, painel:dataPainel, estoque:dataEstoque, movs:dataMovs, vencidas:dataVencidas, clientes:dataClientes, cadastros:dataCadastros};
  el.innerHTML = S.mode==="loading" ? `<div class="empty">Carregando…</div>` : datas[S.tab]();
  ["btnNova","btnRetirar"].forEach(id=>$("#"+id).disabled=S.readonly);
  ["btnSair","btnSair2","ibSair"].forEach(i=>{ const b=$("#"+i); if(b) b.hidden=S.mode!=="db"; });
  const ac=$("#accChip"); if(ac){ ac.textContent=S.mode==="db"?ROLE_LABEL[S.role]+(S.isOwner?" · proprietária":""):""; ac.hidden=S.mode!=="db"; }
}
document.addEventListener("focusout",()=>{ setTimeout(()=>{ if(pendingRefresh) refresh(); },0); });


/* ======================= Evidências (fotos das saídas) ======================= */
const CATS_SAIDA=["Análises","Degustação","Laboratório externo","Carregamento","Cliente","Outras saídas"];
const CATS=[...CATS_SAIDA,"Descarte"];
const CAT_LABEL={"Análises":"Saída para análises","Degustação":"Saída para degustação","Laboratório externo":"Saída para laboratório externo","Carregamento":"Saída para carregamento","Cliente":"Saída para cliente","Outras saídas":"Outras saídas","Descarte":"Saída para descarte"};
const CAT_COLOR={"Análises":"var(--accent)","Degustação":"#d98fc2","Laboratório externo":"var(--ok)","Carregamento":"#7fb0e3","Cliente":"#e0a15f","Outras saídas":"var(--idle)","Descarte":"var(--bad)"};
function catOf(m){ if(m.categoria) return m.categoria; if(m.tipo==="descarte") return "Descarte";
  const t=norm([m.motivo,m.destino].join(" ")); if(/degust/.test(t)) return "Degustação"; if(/extern|laborat/.test(t)) return "Laboratório externo"; if(/carreg/.test(t)) return "Carregamento";
  if(/cliente/.test(t)||m.clienteId) return "Cliente"; if(/analise|estabilidade|auditoria|fiscaliz/.test(t)) return "Análises"; return "Outras saídas"; }
const destOf=m=>m.cliente||(m.destino&&m.destino!=="—"?m.destino:"")||"";

/* fotos: reduz, carimba e guarda no banco (sem precisar de plano pago) */
async function loadImg(file){ try{ return await createImageBitmap(file,{imageOrientation:"from-image"}); }
  catch(e){ return await new Promise((res,rej)=>{ const i=new Image(); i.onload=()=>res(i); i.onerror=rej; i.src=URL.createObjectURL(file); }); } }
function stampTxt(m,etapa){ const d=new Date(); const dt=d.toLocaleString("pt-BR",{dateStyle:"short",timeStyle:"short"});
  return `${m.produto} · lote ${m.lote}${m.ids?" · ID "+m.ids:""}\n${CAT_LABEL[catOf(m)]||catOf(m)} · ${etapa==="preparo"?"preparo":m.tipo==="descarte"?"descarte":"saída"} · ${dt}${m.responsavel?" · "+m.responsavel:""}`; }
async function fotoProc(file,stamp){
  const img=await loadImg(file); const W0=img.width, H0=img.height;
  const draw=(max,q,withStamp)=>{ const k=Math.min(1,max/Math.max(W0,H0)); const w=Math.round(W0*k), h=Math.round(H0*k);
    const c=document.createElement("canvas"); c.width=w; c.height=h; const x=c.getContext("2d"); x.drawImage(img,0,0,w,h);
    if(withStamp&&stamp){ const fs=Math.max(13,Math.round(Math.min(w,h)/34)); x.font=`600 ${fs}px Inter, Arial, sans-serif`; const L=stamp.split("\n");
      const bh=fs*1.4*L.length+fs*.7; x.fillStyle="rgba(0,0,0,.6)"; x.fillRect(0,h-bh,w,bh); x.fillStyle="#fff";
      L.forEach((l,i)=>{ let t=l; while(x.measureText(t).width>w-fs*1.2&&t.length>8) t=t.slice(0,-2); if(t!==l) t=t.slice(0,-1)+"…"; x.fillText(t,fs*.6,h-bh+fs*1.15+i*fs*1.4); }); }
    return c.toDataURL("image/jpeg",q); };
  let q=.8, max=1600, full=draw(max,q,true);
  while(full.length>800000&&q>.4){ q-=.1; if(q<.6) max=1280; full=draw(max,q,true); }
  return {full, thumb:draw(360,.62,false)};
}
async function contarFotos(movId){ try{ const q=await db.collection("evidencias").where("movId","==",movId).get(); await db.collection("movs").doc(movId).update({fotos:q.size}); return q.size; }catch(e){ console.error(e); return null; } }
async function salvarFotos(m,lista){
  if(!lista.length) return 0; if(!db){ toast("Modo demonstração: as fotos não são salvas."); return 0; }
  let n=0;
  for(const f of lista){ try{ const {full,thumb}=await fotoProc(f.file,stampTxt(m,f.etapa));
      await db.collection("evidencias").doc(uid()).set({movId:m.id,etapa:f.etapa,img:full,thumb,produto:m.produto||"",lote:m.lote||"",categoria:catOf(m),data:m.data||"",criadoEm:new Date().toISOString(),por:S.myId||"",nomeArquivo:f.file.name||""});
      n++; }catch(e){ console.error(e); } finally{ try{ URL.revokeObjectURL(f.url); }catch(_){ } } }
  await contarFotos(m.id); return n;
}
/* seletor de fotos usado nos formulários */
function fotoBox(etapas,titulo){ return `<div class="field full"><label>${titulo||"Evidências (fotos)"}</label><div class="fotobox">${etapas.map(([e,l])=>`<div class="fgrp" data-etapa="${e}"><div class="fgh"><b>${l}</b><span class="sub fcnt"></span></div><div class="fthumbs"></div>
  <div class="macts"><label class="btn sm">📷 Câmera<input type="file" accept="image/*" capture="environment" hidden data-etapa="${e}"></label><label class="btn sm">Galeria<input type="file" accept="image/*" multiple hidden data-etapa="${e}"></label></div></div>`).join("")}</div>
  <div class="sub" style="margin-top:6px">Cada foto recebe um carimbo com produto, lote, tipo de saída, data, hora e responsável.</div></div>`; }
function bindFotoBox(root){ const files=[];
  const paint=()=>root.querySelectorAll(".fgrp").forEach(g=>{ const e=g.dataset.etapa; const L=files.map((f,k)=>({...f,k})).filter(f=>f.etapa===e);
    g.querySelector(".fthumbs").innerHTML=L.map(f=>`<span class="fth"><img src="${f.url}" alt="Foto escolhida"><button type="button" class="fx" data-k="${f.k}" aria-label="Remover foto">×</button></span>`).join("");
    g.querySelector(".fcnt").textContent=L.length?`${L.length} foto${L.length===1?"":"s"}`:""; });
  root.addEventListener("change",ev=>{ const inp=ev.target; if(!inp||inp.type!=="file"||!inp.dataset.etapa) return;
    [...inp.files].filter(f=>/^image\//.test(f.type)||/\.(jpe?g|png|heic|webp)$/i.test(f.name)).forEach(f=>files.push({file:f,etapa:inp.dataset.etapa,url:URL.createObjectURL(f)})); inp.value=""; paint(); });
  root.addEventListener("click",ev=>{ const x=ev.target.closest(".fx"); if(!x) return; ev.preventDefault(); const f=files.splice(+x.dataset.k,1)[0]; try{ URL.revokeObjectURL(f.url); }catch(_){ } paint(); });
  files.repaint=paint; return files; }

/* aba Evidências */
function shellEvid(){ const f=S.fe;
  return ptabs("fe","cat",[["","Todas"],...CATS.map(c=>[c,c==="Laboratório externo"?"Lab. externo":c])])+`<div class="toolbar">
    <div class="field grow"><label for="eq">Buscar</label><input id="eq" type="search" placeholder="Produto, lote, ID, cliente, observação" value="${esc(f.q)}"></div>
    <div class="field dsk"><label for="ecat">Tipo de saída</label><select id="ecat"><option value="">Todas</option>${CATS.map(c=>`<option value="${c}"${f.cat===c?" selected":""}>${CAT_LABEL[c]}</option>`).join("")}</select></div>
    <div class="field"><label for="efo">Fotos</label><select id="efo">${[["","Todas"],["sem","Sem fotos"],["com","Com fotos"]].map(([v,l])=>`<option value="${v}"${f.foto===v?" selected":""}>${l}</option>`).join("")}</select></div>
    <div class="field"><label for="ede">De</label><input id="ede" type="date" value="${esc(f.de)}"></div>
    <div class="field"><label for="eat">Até</label><input id="eat" type="date" value="${esc(f.ate)}"></div>
  </div><div id="data"></div>`; }
function evMovs(){ const f=S.fe, q=norm(f.q);
  return S.movs.filter(m=>m.tipo==="saida"||m.tipo==="descarte").map(m=>({...m,cat:catOf(m)}))
    .filter(m=>(!f.de||m.data>=f.de)&&(!f.ate||m.data<=f.ate)&&(!q||norm([m.produto,m.lote,m.ids,m.obs,m.cliente,m.destino,m.motivo,m.responsavel].join(" ")).includes(q)))
    .sort((a,b)=>String(b.data).localeCompare(String(a.data))||String(b.criadoEm).localeCompare(String(a.criadoEm))); }
function dataEvid(){
  const all=evMovs(), f=S.fe;
  const byCat=Object.fromEntries(CATS.map(c=>[c,all.filter(m=>m.cat===c)]));
  const L=all.filter(m=>(!f.cat||m.cat===f.cat)&&(!f.foto||(f.foto==="com"?m.fotos>0:!(m.fotos>0))));
  const base=f.cat?byCat[f.cat]:all; const com=base.filter(m=>m.fotos>0).length;
  const head=`<div class="kpis">
    <div class="kpi" style="--c:var(--accent)"><span class="l">${f.cat?esc(CAT_LABEL[f.cat]):"Saídas registradas"}</span><span class="v">${base.length}</span><span class="s">${f.de||f.ate?"no período escolhido":"desde o início"}</span></div>
    <div class="kpi" style="--c:var(--ok)"><span class="l">Com fotos</span><span class="v">${com}</span><span class="s">${base.length?Math.round(com*100/base.length):0}% das saídas</span></div>
    <div class="kpi" style="--c:var(--warn)"><span class="l">Sem fotos</span><span class="v">${base.length-com}</span><span class="s">dá para incluir depois</span></div>
    <div class="kpi" style="--c:var(--accent)"><span class="l">Fotos guardadas</span><span class="v">${nf(base.reduce((s,m)=>s+(Number(m.fotos)||0),0))}</span><span class="s">saída e preparo</span></div>
  </div>`;
  const grupos=isMobile()?"":`<div class="chips">${CATS.map(c=>`<button type="button" class="chip" data-act="pt" data-o="fe" data-k="cat" data-v="${f.cat===c?"":c}" style="cursor:pointer;border-color:${f.cat===c?CAT_COLOR[c]:"var(--line)"}"><span class="catb" style="--c:${CAT_COLOR[c]}">${esc(c)}</span> <b>${byCat[c].length}</b></button>`).join("")}</div>`;
  if(!all.length) return head+`<div class="panel"><div class="empty"><strong>Nenhuma saída registrada ainda</strong><span>Ao registrar uma retirada ou um descarte, escolha o tipo de saída e anexe as fotos. Elas aparecem aqui organizadas por grupo.</span></div></div>`;
  if(!L.length) return head+grupos+`<div class="panel"><div class="empty"><strong>Nada encontrado nesse filtro</strong></div></div>`;
  const lim=MOBLIM.evid, ro=S.readonly;
  return head+grupos+`<div class="sub">${L.length} saída${L.length===1?"":"s"}</div><div class="mlist cards2">${L.slice(0,lim).map(m=>{ const c=CAT_COLOR[m.cat]; const d=destOf(m);
    return `<div class="mcard" style="--c:${c}">
      <div class="mrow"><div><span class="catb" style="--c:${c}">${esc(CAT_LABEL[m.cat])}</span><div class="mtitle" style="margin-top:6px">${esc(m.produto)}</div><div class="mmeta">Lote <span class="mono">${esc(m.lote)}</span>${m.ids?" · ID "+esc(m.ids):""}</div></div>
        <div style="text-align:right"><div class="mbig">${nf(m.qtd)}</div><div class="mmeta">${m.kg?fkg(m.kg)+" kg":"un"}</div></div></div>
      <div class="mmeta"><span class="mono">${fmtD(m.data)}</span>${m.responsavel?" · "+esc(m.responsavel):""}${d?" · → "+esc(d):""}</div>
      ${m.obs?`<div class="mmeta" style="font-style:italic">${esc(m.obs)}</div>`:""}
      <div class="mrow" style="align-items:center"><span class="fstat" style="color:${m.fotos?"var(--ok)":"var(--warn)"}">${m.fotos?`📷 ${m.fotos} foto${m.fotos===1?"":"s"}`:"Sem fotos"}</span>
        <button class="btn sm ${!m.fotos&&!ro?"primary":""}" data-act="evid" data-id="${m.id}">${m.fotos?"Ver evidências":ro?"Ver":"+ Adicionar fotos"}</button></div>
    </div>`; }).join("")}</div>${L.length>lim?`<button class="btn" style="width:100%" data-act="mais" data-k="evid">Mostrar mais (${L.length-lim} restantes)</button>`:""}`;
}
const EV_CACHE={};
async function evidDetalhe(id){
  const m0=S.movs.find(x=>x.id===id); if(!m0) return; const m={...m0,cat:catOf(m0)}; const d=destOf(m);
  const etapas=m.tipo==="descarte"?[["saida","Fotos do descarte"]]:[["saida","Fotos da saída"],["preparo","Fotos do preparo"]];
  const ro=S.readonly;
  const body=`<div class="ref full"><div><span class="catb" style="--c:${CAT_COLOR[m.cat]}">${esc(CAT_LABEL[m.cat])}</span><div class="prod" style="margin-top:6px">${esc(m.produto)}</div>
     <div class="sub">Lote <span class="lote">${esc(m.lote)}</span> · ${fmtD(m.data)} · ${nf(m.qtd)} un${m.kg?" · "+fkg(m.kg)+" kg":""}${m.ids?" · ID "+esc(m.ids):""}</div>
     <div class="sub">${m.responsavel?esc(m.responsavel):""}${d?" · → "+esc(d):""}${m.motivo?" · "+esc(m.motivo):""}</div>${m.obs?`<div class="sub" style="font-style:italic">${esc(m.obs)}</div>`:""}</div></div>
   <div class="full" id="ev_gal"><p class="muted" style="margin:0">Carregando fotos…</p></div>
   ${ro?"":fotoBox(etapas,"Adicionar fotos")}<div class="formerr full" id="ev_err"></div>`;
  const md=modal("Evidências da saída",body,`<button type="button" class="btn" data-close>Fechar</button>${ro?"":`<button type="submit" class="btn primary">Salvar fotos</button>`}`,true);
  const files=ro?[]:bindFotoBox(md.ov);
  const gal=md.ov.querySelector("#ev_gal");
  const load=async()=>{ if(!db){ gal.innerHTML=`<p class="muted" style="margin:0">Modo demonstração: sem fotos salvas.</p>`; return; }
    let L=[]; try{ const q=await db.collection("evidencias").where("movId","==",id).get(); L=q.docs.map(x=>({id:x.id,...x.data()})).sort((a,b)=>String(a.criadoEm).localeCompare(String(b.criadoEm))); }
    catch(e){ console.error(e); gal.innerHTML=`<p style="margin:0;color:var(--bad)">Não foi possível carregar as fotos. Confira a internet e se as regras do Firebase foram atualizadas.</p>`; return; }
    EV_CACHE[id]=L;
    gal.innerHTML=etapas.map(([e,l])=>{ const G=L.filter(x=>(x.etapa||"saida")===e);
      return `<div class="evg"><div class="fgh"><b>${l}</b><span class="sub">${G.length?G.length+" foto"+(G.length===1?"":"s"):"nenhuma"}</span></div>
      ${G.length?`<div class="evgrid">${G.map(x=>`<button type="button" class="evth" data-ev="${x.id}"><img src="${x.thumb||x.img}" alt="${esc(l)}" loading="lazy"><span>${new Date(x.criadoEm).toLocaleString("pt-BR",{dateStyle:"short",timeStyle:"short"})}</span></button>`).join("")}</div>`:`<p class="sub" style="margin:0">Sem fotos.</p>`}</div>`; }).join(""); };
  load();
  md.ov.addEventListener("click",ev=>{ const t=ev.target.closest(".evth"); if(!t) return; ev.preventDefault(); const x=(EV_CACHE[id]||[]).find(z=>z.id===t.dataset.ev); if(x) verFoto(x,m,load); });
  md.form.onsubmit=async e=>{ e.preventDefault(); if(ro) return; const er=md.ov.querySelector("#ev_err");
    if(!files.length){ er.textContent="Escolha ao menos uma foto (câmera ou galeria)."; return; } er.textContent="";
    const btn=md.form.querySelector('button[type=submit]'); btn.disabled=true; btn.textContent="Salvando…";
    const tot=files.length; const n=await salvarFotos(m,files.splice(0)); files.repaint();
    toast(n===tot?`${n} foto${n===1?"":"s"} salva${n===1?"":"s"}`:`Só ${n} de ${tot} fotos foram salvas.`);
    btn.disabled=false; btn.textContent="Salvar fotos"; load(); };
}
function verFoto(x,m,reload){
  const ov=document.createElement("div"); ov.className="overlay lbx"; const ro=S.readonly;
  const quando=new Date(x.criadoEm).toLocaleString("pt-BR",{dateStyle:"short",timeStyle:"short"});
  ov.innerHTML=`<div class="lb" role="dialog" aria-modal="true" aria-label="Foto de evidência"><img src="${x.img}" alt="Foto de evidência de ${esc(m.produto)}, lote ${esc(m.lote)}">
    <div class="lbbar"><span class="sub">${x.etapa==="preparo"?"Preparo":m.tipo==="descarte"?"Descarte":"Saída"} · ${esc(m.produto)} · lote ${esc(m.lote)} · enviada em ${quando}${x.por?" por "+esc(x.por):""}</span>
    <button type="button" class="btn sm" id="lbDl">Baixar foto</button>${FULL()&&!ro?`<button type="button" class="btn sm danger" id="lbDel">Excluir</button>`:""}<button type="button" class="btn sm primary" id="lbX">Fechar</button></div></div>`;
  document.body.appendChild(ov);
  const close=()=>{ ov.remove(); document.removeEventListener("keydown",k); }; const k=e=>{ if(e.key==="Escape") close(); };
  document.addEventListener("keydown",k); ov.addEventListener("mousedown",e=>{ if(e.target===ov) close(); });
  ov.querySelector("#lbX").onclick=close;
  ov.querySelector("#lbDl").onclick=async()=>{ const b=await (await fetch(x.img)).blob(); await DL.save({filename:`evidencia_${String(m.lote).replace(/[^\w-]+/g,"-")}_${m.data||""}_${x.etapa||"saida"}_${x.id.slice(0,5)}.jpg`,data:b}); };
  const del=ov.querySelector("#lbDel"); if(del) del.onclick=()=>confirmar("Excluir foto","A foto de evidência será apagada definitivamente.","Excluir",async()=>{ await guard(async()=>{ await db.collection("evidencias").doc(x.id).delete(); await contarFotos(m.id); },"Foto excluída"); close(); reload&&reload(); });
}

/* ---------- Mobile (estilo app) ---------- */
const MOBLIM={estoque:40,movs:40,evid:40};
const isMobile=()=>window.matchMedia("(max-width:859px)").matches;
const ST_COLOR={ok:"var(--ok)",avencer:"var(--warn)",vencida:"var(--bad)",esgotada:"var(--idle)",descartada:"var(--idle)"};
function ptabs(obj,key,list){ return `<div class="ptabs" role="group">${list.map(([v,l])=>`<button type="button" class="ptab" data-act="pt" data-o="${obj}" data-k="${key}" data-v="${v}" aria-pressed="${S[obj][key]===v}">${l}</button>`).join("")}</div>`; }
function instalarCard(){ return INSTALL_EVT?`<div class="panel" style="flex-direction:row;align-items:center;gap:12px;flex-wrap:wrap"><img src="icones/icone-192.png" alt="" width="44" height="44" style="border-radius:10px"><div style="flex:1;min-width:180px"><b>Instale o app neste aparelho</b><div class="sub">Abre em tela cheia, com ícone próprio, como um aplicativo.</div></div><button class="btn primary" data-act="instalar">Instalar</button></div>`:""; }
function mobQuick(){ return S.readonly?"":`<div class="mobq"><button class="btn primary" data-act="qnova">+ Nova amostra</button><button class="btn" data-act="qret">Registrar retirada</button></div>`; }
function diasTxt(r){ if(r.dias==null||r.descartada||r.saldo<=0) return ""; return r.dias<=0?`vencida há ${-r.dias} d`:`${r.dias} dias`; }
function cardAmostra(r,ro){
  const kg=r.saldo>0&&!r.descartada?kgOf(r):0;
  return `<div class="mcard" style="--c:${ST_COLOR[r.st]||"var(--accent)"}">
    <div class="mrow"><div><div class="mtitle">${esc(r.produto)}</div><div class="mmeta">Lote <span class="mono">${esc(r.lote)}</span>${r.tipo?" · "+esc(r.tipo):""}</div></div>${pill(r.st)}</div>
    <div class="mrow"><div class="mmeta">Validade <span class="mono">${fmtD(r.validade)}</span>${diasTxt(r)?` · <b style="color:${ST_COLOR[r.st]}">${diasTxt(r)}</b>`:""}<br>${esc(r.local||"sem local")}</div>
      <div style="text-align:right"><div class="mbig">${nf(r.saldo)} <span class="mmeta">${esc(r.unidade||"un")}</span></div><div class="mmeta">${kg?fkg(kg)+" kg":""}</div></div></div>
    <div class="macts">
      ${!ro&&r.st==="vencida"&&r.saldo>0?`<button class="btn sm danger" data-act="descartar" data-id="${r.id}">Descartar</button>`:""}
      ${!ro&&r.saldo>0&&!r.descartada&&r.st!=="vencida"?`<button class="btn sm primary" data-act="retirar" data-id="${r.id}">Retirar</button>`:""}
      ${!ro&&!r.descartada?`<button class="btn sm" data-act="entrada" data-id="${r.id}">+ Entrada</button>`:""}
      ${FULL()&&!ro&&!r.descartada&&idsOf(r).conf?`<button class="btn sm" data-act="ajustar" data-id="${r.id}" style="color:var(--warn);border-color:var(--warn)">Conferir</button>`:""}
      <button class="btn sm" data-act="hist" data-id="${r.id}">Detalhes</button>
    </div></div>`;
}

/* ---------- Painel ---------- */
function shellPainel(){ return `<div id="data"></div>`; }
function dataPainel(){
  const R=rows(); const ativas=R.filter(r=>r.saldo>0&&!r.descartada);
  const un=ativas.reduce((s,r)=>s+r.saldo,0);
  const av=ativas.filter(r=>r.st==="avencer"), vc=ativas.filter(r=>r.st==="vencida");
  const prods=new Set(ativas.map(r=>norm(r.produto))).size;
  if(!S.amostras.length){
    return kpis(0,0,0,0,0) + `<div class="panel"><div class="empty"><strong>Nenhuma amostra cadastrada ainda</strong>
      <span>Cadastre a primeira amostra de produto acabado ou importe a sua planilha atual (aba Cadastros → Importar planilha).</span>
      <div class="actions"><button class="btn primary" data-act="nova">+ Nova amostra</button><button class="btn" data-act="gotab" data-tab="cadastros">Importar planilha</button></div></div></div>`;
  }
  const atn=[...vc,...av].sort((a,b)=>(a.dias??0)-(b.dias??0)).slice(0,10);
  const saidas=saidasAll();
  const hn=S.hist.reduce((s,x)=>s+(Number(x.qtd)||0),0);
  const hnote=hn?`<p class="sub" style="margin:0">Inclui ${nf(hn)} unidades retiradas antes do app (histórico mensal da planilha, só dos lotes válidos importados).</p>`:"";
  return instalarCard() + mobQuick() + kpis(ativas.length,un,prods,av.length,vc.length,ativas.reduce((s,r)=>s+kgOf(r),0)) + `
  <div class="grid2">
    <div class="panel">
      <h2>Retiradas por mês <small>unidades retiradas para análise e envios, últimos 12 meses</small></h2>
      <div class="chart">${chartMeses(saidas)}</div>
      ${medias(saidas)}
      ${hnote}
    </div>
    <div class="panel">
      <h2>Precisa de atenção <small>alerta com ${esc(S.cfg.alerta)} dias</small></h2>
      ${atn.length? `<div class="alist">${atn.map(r=>`<div class="it"><div style="min-width:0"><div class="prod">${esc(r.produto)}</div><div class="sub">Lote <span style="font-family:var(--mono)">${esc(r.lote)}</span> · ${esc(r.local||"sem local")} · saldo ${nf(r.saldo)} ${esc(r.unidade||"")}</div></div>
        <span class="dias" style="color:${r.st==="vencida"?"var(--bad)":"var(--warn)"}">${r.dias<=0?`vencida há ${-r.dias} d`:`${r.dias} d`}</span>
        ${r.st==="vencida"?`<button class="btn sm danger" data-act="descartar" data-id="${r.id}">Descartar</button>`:`<button class="btn sm" data-act="retirar" data-id="${r.id}">Retirar</button>`}</div>`).join("")}</div>`
        : `<div class="empty"><strong>Tudo dentro da validade</strong><span>Nenhuma amostra vence nos próximos ${esc(S.cfg.alerta)} dias.</span></div>`}
    </div>
  </div>
  <div class="panel">
    <h2>Consumo médio por produto <small>médias contadas a partir da primeira retirada de cada produto</small></h2>
    ${consumoProduto(saidas)}
  </div>`;
}
function kpis(a,u,p,av,vc,kg){
  return `<div class="kpis">
    <div class="kpi" style="--c:var(--accent)"><span class="l">Lotes em estoque</span><span class="v">${a}</span><span class="s">${p} produto${p===1?"":"s"} diferente${p===1?"":"s"}</span></div>
    <div class="kpi" style="--c:var(--accent)"><span class="l">Unidades guardadas</span><span class="v">${nf(u)}</span><span class="s">${fkg(kg||0)} kg em estoque</span></div>
    <div class="kpi" style="--c:var(--warn)"><span class="l">A vencer</span><span class="v">${av}</span><span class="s">em até ${esc(S.cfg.alerta)} dias</span></div>
    <div class="kpi" style="--c:var(--bad)"><span class="l">Vencidas</span><span class="v">${vc}</span><span class="s">aguardando descarte</span></div>
  </div>`;
}
function chartMeses(saidas){
  const now=new Date(); const ms=[];
  for(let i=11;i>=0;i--){ const d=new Date(now.getFullYear(),now.getMonth()-i,1); ms.push({k:`${d.getFullYear()}-${pad(d.getMonth()+1)}`,l:MES[d.getMonth()],y:d.getFullYear(),v:0}); }
  for(const m of saidas){ const k=(m.data||"").slice(0,7); const x=ms.find(z=>z.k===k); if(x) x.v+=Number(m.qtd)||0; }
  const max=Math.max(...ms.map(m=>m.v)); const top = max<=0?4:niceMax(max);
  const W=640,H=220,L=40,B=28,T=12,Rr=8; const cw=(W-L-Rr)/12; const h=H-B-T;
  let g=""; for(let i=0;i<=4;i++){ const v=top*i/4; const y=T+h-h*i/4; g+=`<line x1="${L}" x2="${W-Rr}" y1="${y}" y2="${y}" stroke="var(--line)" stroke-width="1"/><text x="${L-6}" y="${y+4}" text-anchor="end">${nf(v)}</text>`; }
  const bars=ms.map((m,i)=>{ const bh=h*m.v/top; const x=L+i*cw+cw*.18; const w=cw*.64; const last=i===11;
    return `<rect x="${x}" y="${T+h-bh}" width="${w}" height="${Math.max(bh,0)}" rx="2" fill="${last?"var(--accent)":"var(--accent-soft)"}" stroke="var(--accent)" stroke-width="${last?0:1}"><title>${m.l}/${m.y}: ${nf(m.v)} un</title></rect>
      ${m.v?`<text x="${x+w/2}" y="${T+h-bh-4}" text-anchor="middle" style="fill:var(--ink)">${nf(m.v)}</text>`:""}
      <text x="${x+w/2}" y="${H-10}" text-anchor="middle">${m.l}</text>`; }).join("");
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Retiradas por mês">${g}${bars}</svg>`;
}
function niceMax(v){ const p=Math.pow(10,Math.floor(Math.log10(v))); const n=v/p; const s=n<=1?1:n<=2?2:n<=2.5?2.5:n<=5?5:10; let t=s*p; if(t<4) t=4; return t; }
function mediasCalc(list){
  if(!list.length) return null;
  const first=list.map(m=>m.data).sort()[0]; const f=parseISO(first), t=new Date();
  const meses=Math.max(1,(t.getFullYear()-f.getFullYear())*12+(t.getMonth()-f.getMonth())+1);
  const tot=list.reduce((s,m)=>s+(Number(m.qtd)||0),0);
  return {tot, m:tot/meses, tri:tot/Math.max(1,Math.ceil(meses/3)), sem:tot/Math.max(1,Math.ceil(meses/6)), ano:tot/Math.max(1,Math.ceil(meses/12)), first};
}
function medias(saidas){
  const c=mediasCalc(saidas);
  if(!c) return `<p class="muted" style="margin:0">As médias aparecem depois da primeira retirada registrada.</p>`;
  return `<div class="avg"><div><span>Média mensal</span><b>${nf(c.m)}</b></div><div><span>Trimestral</span><b>${nf(c.tri)}</b></div><div><span>Semestral</span><b>${nf(c.sem)}</b></div><div><span>Anual</span><b>${nf(c.ano)}</b></div></div>
   <p class="sub" style="margin:0">${nf(c.tot)} unidades retiradas desde ${fmtD(c.first)}.</p>`;
}
function consumoProduto(saidas){
  const by={}; for(const m of saidas){ const k=m.produto||"—"; (by[k]=by[k]||[]).push(m); }
  const R=rows(); const sd={}; for(const r of R){ if(r.saldo>0&&!r.descartada){ sd[r.produto]=(sd[r.produto]||0)+r.saldo; } }
  const names=[...new Set([...Object.keys(by),...Object.keys(sd)])].sort((a,b)=>a.localeCompare(b,"pt"));
  if(!names.length) return `<p class="muted" style="margin:0">Sem dados ainda.</p>`;
  return `<div class="tablebox"><table><thead><tr><th>Produto</th><th class="num">Saldo atual</th><th class="num">Total retirado</th><th class="num">Mensal</th><th class="num">Trimestral</th><th class="num">Semestral</th><th class="num">Anual</th><th class="num">Cobertura</th></tr></thead><tbody>
   ${names.map(n=>{ const c=mediasCalc(by[n]||[]); const s=sd[n]||0; const cob=c&&c.m>0? s/c.m : null;
     return `<tr class="row"><td class="prod">${esc(n)}</td><td class="num">${nf(s)}</td><td class="num">${c?nf(c.tot):"0"}</td><td class="num">${c?nf(c.m):"—"}</td><td class="num">${c?nf(c.tri):"—"}</td><td class="num">${c?nf(c.sem):"—"}</td><td class="num">${c?nf(c.ano):"—"}</td><td class="num">${cob==null?"—":nf(cob)+" meses"}</td></tr>`; }).join("")}
   </tbody></table></div><p class="sub" style="margin:0">Cobertura = saldo atual dividido pela média mensal de retiradas.</p>`;
}

/* ---------- Estoque ---------- */
function opts(list, sel, blank){ return (blank!=null?`<option value="">${esc(blank)}</option>`:"")+list.map(v=>`<option${v===sel?" selected":""}>${esc(v)}</option>`).join(""); }
function shellEstoque(){
  const f=S.f; const locais=[...new Set([...S.cfg.locais,...S.amostras.map(a=>a.local).filter(Boolean)])];
  return mobQuick()+ptabs("f","status",[["ativas","Em estoque"],["avencer","A vencer"],["vencida","Vencidas"],["esgotada","Esgotadas"],["todas","Todas"]])+`<div class="toolbar">
    <div class="field grow"><label for="fq">Buscar</label><input id="fq" type="search" placeholder="Produto, lote, ID (ex.: B14), registro MAPA" value="${esc(f.q)}"></div>
    <div class="field dsk"><label for="fst">Situação</label><select id="fst">
      ${[["ativas","Em estoque"],["ok","Ok"],["avencer","A vencer"],["vencida","Vencidas"],["esgotada","Esgotadas"],["descartada","Descartadas"],["todas","Todas"]].map(([v,l])=>`<option value="${v}"${f.status===v?" selected":""}>${l}</option>`).join("")}</select></div>
    <div class="field dsk"><label for="floc">Localização</label><select id="floc">${opts(locais,f.local,"Todas")}</select></div>
    <div class="field dsk"><label for="ftp">Tipo</label><select id="ftp">${opts(S.cfg.tipos,f.tipo,"Todos")}</select></div>
    <div class="field"><label for="fso">Ordenar por</label><select id="fso">
      ${[["validade","Validade (mais próxima)"],["produto","Produto (A–Z)"],["fabricacao","Fabricação (recente)"],["saldo","Saldo (maior)"]].map(([v,l])=>`<option value="${v}"${f.sort===v?" selected":""}>${l}</option>`).join("")}</select></div>
  </div><div id="data"></div>`;
}
function filtraEstoque(){
  const f=S.f, q=norm(f.q);
  let R=rows().filter(r=>{
    if(f.status==="ativas" && (r.saldo<=0||r.descartada)) return false;
    if(!["ativas","todas"].includes(f.status) && r.st!==f.status) return false;
    if(f.local && r.local!==f.local) return false;
    if(f.tipo && r.tipo!==f.tipo) return false;
    if(q && !norm([r.produto,r.lote,r.codigo,r.mapa,r.obs,r.local,idsOf(r).ids.join(" ")].join(" ")).includes(q)) return false;
    return true; });
  const s=f.sort;
  R.sort((a,b)=> s==="produto"? a.produto.localeCompare(b.produto,"pt")||String(a.lote).localeCompare(String(b.lote))
    : s==="fabricacao"? String(b.fabricacao||"").localeCompare(String(a.fabricacao||""))
    : s==="saldo"? b.saldo-a.saldo
    : (ST_ORDER[a.st]-ST_ORDER[b.st]) || String(a.validade||"9").localeCompare(String(b.validade||"9")));
  return R;
}
function dataEstoque(){
  const R=filtraEstoque();
  if(!S.amostras.length) return `<div class="panel"><div class="empty"><strong>Nenhuma amostra cadastrada</strong><span>Use “Nova amostra” para registrar o primeiro lote guardado.</span><button class="btn primary" data-act="nova">+ Nova amostra</button></div></div>`;
  if(!R.length) return `<div class="panel"><div class="empty"><strong>Nenhuma amostra com esses filtros</strong><span>Troque a situação para “Todas” ou limpe a busca.</span></div></div>`;
  const ro=S.readonly;
  const resumo=`<div class="sub">${R.length} lote${R.length===1?"":"s"} · ${nf(R.reduce((s,r)=>s+Math.max(r.saldo,0),0))} unidades · ${fkg(R.reduce((s,r)=>s+(r.saldo>0&&!r.descartada?kgOf(r):0),0))} kg</div>`;
  if(isMobile()){ const lim=MOBLIM.estoque; return resumo+`<div class="mlist">${R.slice(0,lim).map(r=>cardAmostra(r,ro)).join("")}</div>${R.length>lim?`<button class="btn" style="width:100%" data-act="mais" data-k="estoque">Mostrar mais (${R.length-lim} restantes)</button>`:""}`; }
  return resumo+`
  <div class="tablebox"><table><thead><tr><th>Produto</th><th>Lote</th><th>Tipo</th><th>Fabricação</th><th>Validade</th><th class="num">Dias</th><th>Situação</th><th class="num">Saldo</th><th class="num">Peso (kg)</th><th>Localização</th><th>Atualização</th><th></th></tr></thead><tbody>
  ${R.map(r=>`<tr class="row">
    <td><div class="prod">${esc(r.produto)}</div><div class="sub">${[r.mapa?"MAPA "+r.mapa:"",r.codigo,r.conteudo].filter(Boolean).map(esc).join(" · ")}</div></td>
    <td class="mono">${esc(r.lote)}</td><td class="nowrap">${esc(r.tipo||"—")}</td>
    <td class="mono">${fmtD(r.fabricacao)}</td><td class="mono">${fmtD(r.validade)}</td>
    <td class="num" style="color:${r.st==="vencida"?"var(--bad)":r.st==="avencer"?"var(--warn)":"inherit"}">${r.dias==null||r.descartada||r.saldo<=0?"—":r.dias}</td>
    <td>${pill(r.st)}</td><td class="num"><b>${nf(r.saldo)}</b> <span class="sub">${esc(r.unidade||"")}</span></td><td class="num">${(k=>k?fkg(k):"—")(r.saldo>0&&!r.descartada?kgOf(r):0)}</td>
    <td class="nowrap">${esc(r.local||"—")}</td><td><div class="mono">${fmtD(r.atual)}</div><div class="sub">${esc(r.respUlt)}</div></td>
    <td><div class="rowacts">
      ${!ro&&r.saldo>0&&!r.descartada?`<button class="btn sm" data-act="retirar" data-id="${r.id}" title="Registrar retirada">Retirar</button>`:""}
      ${!ro&&!r.descartada?`<button class="btn sm" data-act="entrada" data-id="${r.id}" title="Adicionar unidades">+ Entrada</button>`:""}
      ${FULL()&&!ro&&!r.descartada&&idsOf(r).conf?`<button class="btn sm" data-act="ajustar" data-id="${r.id}" title="IDs não batem com o saldo" style="color:var(--warn);border-color:var(--warn)">Conferir</button>`:""}
      <button class="btn sm" data-act="hist" data-id="${r.id}" title="Histórico e edição">Detalhes</button>
    </div></td></tr>`).join("")}
  </tbody></table></div>`;
}

/* ---------- Movimentações ---------- */
function shellMovs(){
  const f=S.fm; const resps=[...new Set([...S.cfg.analistas,...S.movs.map(m=>m.responsavel).filter(Boolean)])].sort((a,b)=>a.localeCompare(b,"pt"));
  return ptabs("fm","tipo",[["","Todas"],["entrada","Entradas"],["saida","Retiradas"],["descarte","Descartes"],["ajuste","Ajustes"]])+`<div class="toolbar">
    <div class="field grow"><label for="mq">Buscar</label><input id="mq" type="search" placeholder="Produto, lote, motivo" value="${esc(f.q)}"></div>
    <div class="field dsk"><label for="mtp">Tipo</label><select id="mtp"><option value="">Todos</option>${Object.entries(MV_LABEL).map(([v,l])=>`<option value="${v}"${f.tipo===v?" selected":""}>${l}</option>`).join("")}</select></div>
    <div class="field"><label for="mrs">Responsável</label><select id="mrs">${opts(resps,f.resp,"Todos")}</select></div>
    <div class="field"><label for="mde">De</label><input id="mde" type="date" value="${esc(f.de)}"></div>
    <div class="field"><label for="mat">Até</label><input id="mat" type="date" value="${esc(f.ate)}"></div>
  </div><div id="data"></div>`;
}
function filtraMovs(){
  const f=S.fm,q=norm(f.q);
  return S.movs.filter(m=>(!f.tipo||m.tipo===f.tipo)&&(!f.resp||m.responsavel===f.resp)&&(!f.de||m.data>=f.de)&&(!f.ate||m.data<=f.ate)
    &&(!q||norm([m.produto,m.lote,m.motivo,m.obs,m.destino].join(" ")).includes(q)))
    .sort((a,b)=>String(b.data).localeCompare(String(a.data))||String(b.criadoEm).localeCompare(String(a.criadoEm)));
}
function dataMovs(){
  const M=filtraMovs();
  if(!S.movs.length) return `<div class="panel"><div class="empty"><strong>Nenhuma movimentação ainda</strong><span>Entradas, retiradas e descartes aparecem aqui assim que forem registrados.</span></div></div>`;
  if(!M.length) return `<div class="panel"><div class="empty"><strong>Nada encontrado nesse filtro</strong></div></div>`;
  const tot=t=>M.filter(m=>m.tipo===t).reduce((s,m)=>s+(Number(m.qtd)||0),0);
  const resumo=`<div class="sub">${M.length} registro${M.length===1?"":"s"} · entradas ${nf(tot("entrada"))} · retiradas ${nf(tot("saida"))} · descartes ${nf(tot("descarte"))}</div>`;
  const sinal=m=>Number(m.qtd)===0?"":(m.tipo==="entrada"||m.tipo==="reclass_in"||m.tipo==="ajuste"&&m.qtd>0?"+":"−");
  if(isMobile()){ const lim=MOBLIM.movs; const MC={entrada:"var(--ok)",reclass_in:"var(--ok)",saida:"var(--accent)",descarte:"var(--bad)",reclass_out:"var(--warn)",ajuste:"var(--idle)"};
    return resumo+`<div class="mlist">${M.slice(0,lim).map(m=>`<div class="mcard" style="--c:${MC[m.tipo]||"var(--accent)"}">
      <div class="mrow"><div><div class="mtitle">${esc(m.produto)}</div><div class="mmeta">Lote <span class="mono">${esc(m.lote)}</span>${m.ids?" · ID "+esc(m.ids):""}</div></div>
        <div style="text-align:right"><div class="mbig" style="color:${MC[m.tipo]||"inherit"}">${sinal(m)}${nf(Math.abs(m.qtd))}</div><div class="mmeta">${m.kg?fkg(m.kg)+" kg":""}</div></div></div>
      <div class="mmeta"><b class="mv-${m.tipo}">${MV_LABEL[m.tipo]||esc(m.tipo)}</b>${m.tipo==="saida"||m.tipo==="descarte"?" · "+esc(catOf(m)):""} · <span class="mono">${fmtD(m.data)}</span>${m.responsavel?" · "+esc(m.responsavel):""}${m.fotos?` · 📷 ${m.fotos}`:""}</div>
      ${m.motivo||m.cliente||m.destino?`<div class="mmeta">${[m.motivo,m.cliente||m.destino].filter(Boolean).map(esc).join(" → ")}</div>`:""}
      ${m.obs?`<div class="mmeta" style="font-style:italic">${esc(m.obs)}</div>`:""}
      ${(m.tipo==="saida"||m.tipo==="descarte")||(FULL()&&!S.readonly)?`<div class="macts">${m.tipo==="saida"||m.tipo==="descarte"?`<button class="btn sm" data-act="evid" data-id="${m.id}">Evidências${m.fotos?` (${m.fotos})`:""}</button>`:""}${!FULL()||S.readonly?"":`<button class="btn sm" data-act="delmov" data-id="${m.id}">Desfazer</button>`}</div>`:""}</div>`).join("")}</div>${M.length>lim?`<button class="btn" style="width:100%" data-act="mais" data-k="movs">Mostrar mais (${M.length-lim} restantes)</button>`:""}`; }
  return resumo+`
  <div class="tablebox"><table><thead><tr><th>Data</th><th>Tipo</th><th>Produto</th><th>Lote</th><th class="num">Qtd</th><th>Motivo</th><th>Destino</th><th>Responsável</th><th>Obs.</th><th></th></tr></thead><tbody>
  ${M.slice(0,500).map(m=>`<tr class="row"><td class="mono">${fmtD(m.data)}</td><td class="nowrap mv-${m.tipo}"><b>${MV_LABEL[m.tipo]||esc(m.tipo)}</b></td>
    <td class="prod">${esc(m.produto)}</td><td class="mono">${esc(m.lote)}${m.ids?`<div class="sub">ID ${esc(m.ids)}</div>`:""}${m.kg?`<div class="sub">${fkg(m.kg)} kg</div>`:""}</td>
    <td class="num">${Number(m.qtd)===0?"":(m.tipo==="entrada"||m.tipo==="reclass_in"||m.tipo==="ajuste"&&m.qtd>0?"+":"−")}${nf(Math.abs(m.qtd))}</td>
    <td>${esc(m.motivo||"")}</td><td>${esc(m.cliente||m.destino||"")}</td><td class="nowrap">${esc(m.responsavel||"")}</td><td class="sub">${esc(m.obs||"")}</td>
    <td><div class="rowacts">${m.tipo==="saida"||m.tipo==="descarte"?`<button class="btn sm" data-act="evid" data-id="${m.id}" title="Fotos de evidência">📷 ${m.fotos||0}</button>`:""}${!FULL()||S.readonly?"":`<button class="btn sm" data-act="delmov" data-id="${m.id}" title="Desfazer lançamento">Desfazer</button>`}</div></td></tr>`).join("")}
  </tbody></table></div>${M.length>500?`<div class="sub">Mostrando os 500 mais recentes. Use os filtros ou exporte para ver todos.</div>`:""}`;
}

/* ---------- Vencidas ---------- */
function shellVencidas(){ return `<div id="data"></div>`; }
function dataVencidas(){
  const R=rows().filter(r=>r.st==="vencida").sort((a,b)=>a.dias-b.dias);
  const desc=S.movs.filter(m=>m.tipo==="descarte").sort((a,b)=>String(b.data).localeCompare(String(a.data))).slice(0,30);
  if(isMobile()) return `<div class="sub">${R.length} lote${R.length===1?"":"s"} vencido${R.length===1?"":"s"} aguardando descarte</div>
    ${R.length?`<div class="mlist">${R.map(r=>cardAmostra(r,S.readonly)).join("")}</div>`:`<div class="panel"><div class="empty"><strong>Nenhuma amostra vencida em estoque</strong></div></div>`}
    <h2 style="margin:8px 0 0;font-size:18px">Descartes recentes</h2>
    ${desc.length?`<div class="mlist">${desc.map(m=>`<div class="mcard" style="--c:var(--bad)"><div class="mrow"><div><div class="mtitle">${esc(m.produto)}</div><div class="mmeta">Lote <span class="mono">${esc(m.lote)}</span> · <span class="mono">${fmtD(m.data)}</span></div><div class="mmeta">${esc(m.motivo||"")}${m.responsavel?" · "+esc(m.responsavel):""}</div></div><div class="mbig">${nf(m.qtd)}</div></div></div>`).join("")}</div>`:`<p class="muted" style="margin:0">Nenhum descarte registrado.</p>`}`;
  return `<div class="panel"><h2>Vencidas aguardando descarte <small>${R.length} lote${R.length===1?"":"s"}</small></h2>
  ${R.length?`<div class="tablebox"><table><thead><tr><th>Produto</th><th>Lote</th><th>Validade</th><th class="num">Vencida há</th><th class="num">Saldo</th><th>Localização</th><th></th></tr></thead><tbody>
   ${R.map(r=>`<tr class="row"><td class="prod">${esc(r.produto)}</td><td class="mono">${esc(r.lote)}</td><td class="mono">${fmtD(r.validade)}</td><td class="num" style="color:var(--bad)">${-r.dias} d</td><td class="num">${nf(r.saldo)} ${esc(r.unidade||"")}</td><td>${esc(r.local||"—")}</td>
   <td><div class="rowacts">${S.readonly?"":`<button class="btn sm danger" data-act="descartar" data-id="${r.id}">Registrar descarte</button>`}</div></td></tr>`).join("")}
   </tbody></table></div>`:`<div class="empty"><strong>Nenhuma amostra vencida em estoque</strong><span>Quando um lote passar da validade ele aparece aqui para o descarte ser registrado.</span></div>`}
  </div>
  <div class="panel"><h2>Descartes recentes <small>últimos 30</small></h2>
  ${desc.length?`<div class="tablebox"><table><thead><tr><th>Data</th><th>Produto</th><th>Lote</th><th class="num">Qtd</th><th>Motivo</th><th>Responsável</th></tr></thead><tbody>
   ${desc.map(m=>`<tr class="row"><td class="mono">${fmtD(m.data)}</td><td class="prod">${esc(m.produto)}</td><td class="mono">${esc(m.lote)}</td><td class="num">${nf(m.qtd)}</td><td>${esc(m.motivo)}</td><td>${esc(m.responsavel)}</td></tr>`).join("")}
   </tbody></table></div>`:`<p class="muted" style="margin:0">Nenhum descarte registrado.</p>`}</div>`;
}

/* ---------- Clientes ---------- */
function enviosAll(){
  const app=S.movs.filter(m=>m.tipo==="saida"&&m.clienteId).map(m=>({...m,origem:"app"}));
  return [...app, ...S.envh.map(e=>({...e,ids:e.ids||S.envIds[e.id]||"",kg:e.kg||S.envKg[e.id]||0}))];
}
function statsClientes(){
  const E=enviosAll(), st={};
  for(const e of E){ const s=st[e.clienteId]=st[e.clienteId]||{n:0,un:0,ult:"",prods:{},lotes:new Set()};
    s.n++; s.un+=Number(e.qtd)||0; if(e.data&&e.data>s.ult) s.ult=e.data; s.prods[e.produto]=(s.prods[e.produto]||0)+(Number(e.qtd)||0); s.lotes.add(e.lote); }
  return st;
}
const isLab = id=>{ const c=S.clientes.find(x=>x.id===id); return !!c&&c.tipo==="Laboratório externo"; };
const PAISES = ()=>[...new Set(S.clientes.map(c=>c.pais).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"pt"));
function shellClientes(){
  const f=S.fc;
  return ptabs("fc","tipo",[["","Todos"],["Cliente","Clientes"],["Laboratório externo","Laboratórios"],["Outro","Outros"]])+`<div class="toolbar">
    <div class="field grow"><label for="cq">Buscar</label><input id="cq" type="search" placeholder="Nome, país, cidade ou contato" value="${esc(f.q)}"></div>
    <div class="field"><label for="cpa">País</label><select id="cpa">${opts(PAISES(),f.pais,"Todos")}</select></div>
    <div class="field dsk"><label for="ctp">Tipo</label><select id="ctp">${opts(["Cliente","Laboratório externo","Outro"],f.tipo,"Todos")}</select></div>
    <button class="btn primary" data-act="novocli"${!FULL()||S.readonly?" disabled":""}>+ Novo cadastro</button>
  </div><div id="data"></div>`;
}
function dataClientes(){
  const st=statsClientes(), f=S.fc, q=norm(f.q);
  const L=S.clientes.filter(c=>(!f.pais||c.pais===f.pais)&&(!f.tipo||(c.tipo||"Cliente")===f.tipo)&&(!q||norm([c.nome,c.pais,c.cidade,c.contato,c.email].join(" ")).includes(q)))
    .map(c=>({...c,s:st[c.id]||{n:0,un:0,ult:"",prods:{},lotes:new Set()}}))
    .sort((a,b)=>String(b.s.ult).localeCompare(String(a.s.ult))||a.nome.localeCompare(b.nome,"pt"));
  if(!S.clientes.length) return `<div class="panel"><div class="empty"><strong>Nenhum cliente ou laboratório cadastrado</strong><span>Cadastre quem recebe amostras. Depois, ao registrar uma retirada com motivo “Envio a cliente” ou “Análise externa”, escolha o destinatário.</span><button class="btn primary" data-act="novocli">+ Novo cadastro</button></div></div>`;
  const E=enviosAll(); const lim=isoLocal(new Date(new Date().getFullYear()-1,new Date().getMonth(),new Date().getDate()));
  const e12=E.filter(e=>e.data&&e.data>=lim); const ativos=new Set(e12.map(e=>e.clienteId)).size;
  const head=`<div class="kpis">
    <div class="kpi" style="--c:var(--accent)"><span class="l">Cadastros</span><span class="v">${S.clientes.length}</span><span class="s">${(n=>n+(n===1?" cliente":" clientes"))(S.clientes.filter(c=>(c.tipo||"Cliente")==="Cliente").length)} · ${(n=>n+(n===1?" laboratório":" laboratórios"))(S.clientes.filter(c=>c.tipo==="Laboratório externo").length)}</span></div>
    <div class="kpi" style="--c:var(--accent)"><span class="l">Atendidos em 12 meses</span><span class="v">${ativos}</span><span class="s">receberam ao menos um envio</span></div>
    <div class="kpi" style="--c:var(--ok)"><span class="l">Envios em 12 meses</span><span class="v">${e12.length}</span><span class="s">lançamentos de produto/lote</span></div>
    <div class="kpi" style="--c:var(--ok)"><span class="l">Unidades enviadas</span><span class="v">${nf(e12.reduce((s,e)=>s+(Number(e.qtd)||0),0))}</span><span class="s">nos últimos 12 meses</span></div>
  </div>`;
  if(!L.length) return head+`<div class="panel"><div class="empty"><strong>Nenhum cadastro nesse filtro</strong></div></div>`;
  const ro=S.readonly;
  if(isMobile()) return head+`<div class="mlist">${L.map(c=>{ const top=Object.entries(c.s.prods).sort((a,b)=>b[1]-a[1]).slice(0,2);
    return `<div class="mcard" style="--c:${c.tipo==="Laboratório externo"?"var(--ok)":"var(--accent)"}">
      <div class="mrow"><div><div class="mtitle">${esc(c.nome)}</div><div class="mmeta">${[c.tipo||"Cliente",c.pais,c.cidade].filter(Boolean).map(esc).join(" · ")}</div></div>
        <div style="text-align:right"><div class="mbig">${c.s.n}</div><div class="mmeta">envios</div></div></div>
      <div class="mmeta">${nf(c.s.un)} unidades · último envio <span class="mono">${fmtD(c.s.ult)}</span></div>
      ${top.length?`<div class="mmeta">${top.map(([p,q])=>esc(p)+" ("+nf(q)+")").join(" · ")}</div>`:""}
      <div class="macts">${ro?"":`<button class="btn sm primary" data-act="enviarcli" data-id="${c.id}">Enviar amostra</button>`}<button class="btn sm" data-act="cli" data-id="${c.id}">Detalhes</button></div></div>`; }).join("")}</div>`;
  return head+`<div class="sub">${L.length} cadastro${L.length===1?"":"s"}${S.envh.length?" · inclui envios da planilha dos lotes válidos importados":""}</div>
  <div class="tablebox"><table><thead><tr><th>Cliente / laboratório</th><th>País</th><th class="num">Envios</th><th class="num">Unidades</th><th>Produtos mais enviados</th><th>Último envio</th><th></th></tr></thead><tbody>
  ${L.map(c=>{ const top=Object.entries(c.s.prods).sort((a,b)=>b[1]-a[1]).slice(0,2);
    return `<tr class="row"><td><div class="prod">${esc(c.nome)}</div><div class="sub">${[c.tipo&&c.tipo!=="Cliente"?c.tipo:"",c.cidade,c.contato].filter(Boolean).map(esc).join(" · ")}</div></td>
    <td class="nowrap">${esc(c.pais||"—")}</td><td class="num">${c.s.n}</td><td class="num">${nf(c.s.un)}</td>
    <td>${top.length?top.map(([p,q])=>`<div class="sub" style="color:var(--ink)">${esc(p)} <span class="muted">(${nf(q)})</span></div>`).join(""):`<span class="sub">—</span>`}</td>
    <td class="mono">${fmtD(c.s.ult)}</td>
    <td><div class="rowacts">${ro?"":`<button class="btn sm" data-act="enviarcli" data-id="${c.id}">Enviar amostra</button>`}<button class="btn sm" data-act="cli" data-id="${c.id}">Detalhes</button></div></td></tr>`; }).join("")}
  </tbody></table></div>`;
}
function formCliente(c){
  if(S.readonly||!FULL()) return; const ed=!!c; c=c||{tipo:"Cliente"};
  const body=`<div class="field full"><label for="k_nome">Nome *</label><input id="k_nome" value="${esc(c.nome)}" required></div>
   <div class="field"><label for="k_tipo">Tipo</label><select id="k_tipo">${opts(["Cliente","Laboratório externo","Outro"],c.tipo||"Cliente")}</select></div>
   <div class="field"><label for="k_pais">País</label><input id="k_pais" list="dlpa" value="${esc(c.pais)}">${dl("dlpa",PAISES())}</div>
   <div class="field"><label for="k_cid">Cidade / estado</label><input id="k_cid" value="${esc(c.cidade)}"></div>
   <div class="field"><label for="k_cont">Contato</label><input id="k_cont" value="${esc(c.contato)}" placeholder="Nome da pessoa"></div>
   <div class="field full"><label for="k_mail">E-mail / telefone</label><input id="k_mail" value="${esc(c.email)}"></div>
   <div class="field full"><label for="k_obs">Observação</label><textarea id="k_obs" placeholder="ex.: exige contraprova de cada lote, envio via trading">${esc(c.obs)}</textarea></div>
   <div class="formerr full" id="k_err"></div>`;
  const m=modal(ed?"Editar cadastro":"Novo cliente / laboratório",body,`${ed?`<button type="button" class="btn danger" id="k_del" style="margin-right:auto">Excluir</button>`:""}<button type="button" class="btn" data-close>Cancelar</button><button type="submit" class="btn primary">${ed?"Salvar":"Cadastrar"}</button>`);
  if(ed) m.ov.querySelector("#k_del").onclick=()=>{ m.close(); const n=(statsClientes()[c.id]||{n:0}).n;
    confirmar("Excluir cadastro", n?`${c.nome} tem ${n} envio(s) registrados. O cadastro some da lista, mas os envios continuam nas movimentações.`:`Excluir ${c.nome}?`,"Excluir",()=>guard(async()=>{ if(db) await db.collection("clientes").doc(c.id).delete(); else { S.clientes=S.clientes.filter(x=>x.id!==c.id); refresh(); } },"Cadastro excluído")); };
  m.form.onsubmit=async e=>{ e.preventDefault(); const g=id=>m.ov.querySelector("#"+id).value.trim();
    const d={nome:g("k_nome"),tipo:g("k_tipo"),pais:g("k_pais"),cidade:g("k_cid"),contato:g("k_cont"),email:g("k_mail"),obs:g("k_obs")};
    if(!d.nome){ m.ov.querySelector("#k_err").textContent="Informe o nome."; return; }
    if(S.clientes.some(x=>x.id!==c.id&&norm(x.nome)===norm(d.nome))){ m.ov.querySelector("#k_err").textContent="Já existe um cadastro com esse nome."; return; }
    m.close();
    await guard(async()=>{ if(ed){ if(db) await db.collection("clientes").doc(c.id).update(d); else { Object.assign(S.clientes.find(x=>x.id===c.id),d); refresh(); } }
      else { const id="c-"+uid(); if(db) await db.collection("clientes").doc(id).set({...d,criadoEm:new Date().toISOString()}); else { S.clientes.push({id,...d}); refresh(); } } }, ed?"Cadastro atualizado":"Cadastro salvo");
  };
}
function clienteDetalhe(id){
  const c=S.clientes.find(x=>x.id===id); if(!c) return;
  const E=enviosAll().filter(e=>e.clienteId===id).sort((a,b)=>String(b.data).localeCompare(String(a.data)));
  const byP={}; for(const e of E) byP[e.produto]=(byP[e.produto]||0)+(Number(e.qtd)||0);
  const R=rows(); const stOf=e=>{ const r=R.find(x=>x.id===e.amostraId); return r?pill(r.st):""; };
  const info=[["Tipo",c.tipo||"Cliente"],["País",c.pais],["Cidade",c.cidade],["Contato",c.contato],["E-mail / telefone",c.email],["Observação",c.obs]].filter(x=>x[1]);
  const body=`<div class="ref full"><div><div class="prod">${esc(c.nome)}</div><div class="sub">${E.length} envio${E.length===1?"":"s"} · ${nf(E.reduce((s,e)=>s+(Number(e.qtd)||0),0))} unidades · ${fkg(E.reduce((s,e)=>s+(Number(e.kg)||0),0))} kg</div></div></div>
   <div class="full" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px">${info.map(([k,v])=>`<div><div class="sub">${k}</div><div>${esc(v)}</div></div>`).join("")}</div>
   ${Object.keys(byP).length?`<div class="full"><div class="sub" style="margin-bottom:4px">Por produto</div><div class="chips">${Object.entries(byP).sort((a,b)=>b[1]-a[1]).map(([p,q])=>`<span class="chip" style="padding-right:10px">${esc(p)} · <b>${nf(q)}</b></span>`).join("")}</div></div>`:""}
   <div class="full tablebox"><table><thead><tr><th>Data</th><th>Produto</th><th>Lote</th><th>ID</th><th class="num">Qtd</th><th class="num">Peso (kg)</th><th>Situação do lote</th><th>Responsável</th><th>Obs.</th></tr></thead><tbody>
   ${E.map(e=>`<tr><td class="mono">${fmtD(e.data)}</td><td>${esc(e.produto)}</td><td class="mono">${esc(e.lote)}</td><td class="mono" style="white-space:normal;min-width:80px">${esc(e.ids||"—")}</td><td class="num">${nf(e.qtd)}</td><td class="num">${e.kg?fkg(e.kg):"—"}</td><td>${stOf(e)}</td><td class="nowrap">${esc(e.responsavel)}</td><td class="sub">${esc(e.obs)}${e.origem==="planilha"?` <span class="muted">(planilha)</span>`:""}</td></tr>`).join("")||`<tr><td colspan="9" class="muted">Nenhum envio registrado.</td></tr>`}
   </tbody></table></div>`;
  const ro=S.readonly;
  const m=modal(c.tipo==="Laboratório externo"?"Laboratório externo":"Cliente",body,`${ro?"":`${FULL()?`<button type="button" class="btn" id="cd_ed" style="margin-right:auto">Editar cadastro</button>`:""}<button type="button" class="btn" id="cd_env">Enviar amostra</button>`}<button type="button" class="btn primary" data-close>Fechar</button>`,true);
  if(!ro){ const ce=m.ov.querySelector("#cd_ed"); if(ce) ce.onclick=()=>{ m.close(); formCliente(c); }; m.ov.querySelector("#cd_env").onclick=()=>{ m.close(); formMov(null,"saida",id); }; }
}

/* ---------- Acessos (só a proprietária) ---------- */
function acessosPanel(){
  const P=S.pessoas.filter(p=>p.id!==S.myId&&!p.origem);
  const pend=P.filter(p=>!S.acessos[p.id]).length;
  const rowsH=P.sort((a,b)=>(!!S.acessos[a.id])-(!!S.acessos[b.id])||String(a.nome||a.id).localeCompare(String(b.nome||b.id),"pt")).map(p=>{ const r=S.acessos[p.id]||"pendente";
    return `<tr class="row"><td><div class="prod">${esc(p.nome||"—")}</div><div class="sub">${esc(p.email||p.id)}</div></td><td class="mono">${p.primeiroAcesso?fmtD(p.primeiroAcesso.slice(0,10)):"—"}</td><td class="mono">${p.ultimoAcesso?fmtD(p.ultimoAcesso.slice(0,10)):"—"}</td>
    <td><select class="accsel" data-uid="${esc(p.id)}" aria-label="Nível de acesso" style="${r==="pendente"?"border-color:var(--warn);color:var(--warn)":""}">${["pendente","leitura","operador","total","bloqueado"].map(x=>`<option value="${x}"${r===x?" selected":""}>${ROLE_LABEL[x]}</option>`).join("")}</select></td>
    <td><div class="rowacts"><button type="button" class="btn sm accreset" data-uid="${esc(p.id)}">Enviar e-mail de nova senha</button></div></td></tr>`; }).join("");
  const sel=(p,r)=>`<select class="accsel" data-uid="${esc(p.id)}" aria-label="Nível de acesso de ${esc(p.nome||p.id)}" style="${r==="pendente"?"border-color:var(--warn);color:var(--warn)":""}">${["pendente","leitura","operador","total","bloqueado"].map(x=>`<option value="${x}"${r===x?" selected":""}>${ROLE_LABEL[x]}</option>`).join("")}</select>`;
  const lista=isMobile()?(P.length?`<div class="mlist">${P.map(p=>{ const r=S.acessos[p.id]||"pendente"; return `<div class="mcard" style="--c:${r==="pendente"?"var(--warn)":r==="bloqueado"?"var(--bad)":"var(--ok)"}"><div><div class="mtitle">${esc(p.nome||"—")}</div><div class="mmeta">${esc(p.email||p.id)}</div><div class="mmeta">Último acesso ${p.ultimoAcesso?fmtD(p.ultimoAcesso.slice(0,10)):"—"}</div></div>${sel(p,r)}<div class="macts"><button type="button" class="btn sm accreset" data-uid="${esc(p.id)}">Enviar e-mail de nova senha</button></div></div>`; }).join("")}</div>`:`<p class="muted" style="margin:0">Ninguém ainda. Envie o endereço do app para a equipe e peça para criarem a conta.</p>`):null;
  return `<div class="panel"><h2>Acessos <small>só a proprietária vê e altera${pend?` · <b style="color:var(--warn)">${pend} aguardando liberação</b>`:""}</small></h2>
    <p class="sub" style="margin:0">Cada pessoa abre o app, toca em <b>Criar conta</b>, confirma o e-mail e aparece aqui como “Aguardando liberação”. Escolha o nível: <b>Acesso total</b> (tudo, inclusive cadastros, ajustes, reclassificação e histórico de acesso) · <b>Inclusão e saída</b> (nova amostra, entrada, retirada, descarte e envio) · <b>Visualização</b> (só consulta) · <b>Sem acesso</b> (não entra). Quem está aguardando liberação não vê nenhum dado.</p>
    ${lista!==null?lista:`<div class="tablebox"><table><thead><tr><th>Pessoa</th><th>Primeiro acesso</th><th>Último acesso</th><th>Nível no app</th><th></th></tr></thead><tbody>
    ${rowsH||`<tr><td colspan="5" class="muted">Ninguém ainda. Envie o endereço do app para a equipe e peça para criarem a conta.</td></tr>`}</tbody></table></div>`}</div>`;
}
/* Histórico de acesso (acesso total) */
let ACC_LOG=null, ACC_F="";
function historicoAcessoPanel(){
  const pessoas=S.pessoas.map(p=>p.id);
  return `<div class="panel"><h2>Histórico de acesso <small>contas criadas, entradas e saídas</small></h2>
    ${ACC_LOG?`<div class="toolbar"><div class="field"><label for="accF">Pessoa</label><select id="accF"><option value="">Todas</option>${pessoas.map(u=>`<option value="${esc(u)}"${ACC_F===u?" selected":""} class="accopt" data-uid="${esc(u)}">…</option>`).join("")}</select></div><button type="button" class="btn" id="accLoad">Atualizar</button></div>
    <div class="tablebox"><table><thead><tr><th>Data e hora</th><th>Pessoa</th><th>Evento</th><th>Aparelho</th></tr></thead><tbody>
    ${ACC_LOG.filter(x=>!ACC_F||x.uid===ACC_F).slice(0,400).map(x=>`<tr class="row"><td class="mono">${new Date(x.t).toLocaleString("pt-BR",{dateStyle:"short",timeStyle:"short"})}</td><td><span class="accname" data-uid="${esc(x.uid)}">…</span></td><td style="color:${x.ev==="falha"||x.ev==="bloqueio"?"var(--bad)":x.ev==="entrada"?"var(--ok)":"inherit"}"><b>${esc(EV_LABEL[x.ev]||x.ev)}</b></td><td class="sub">${esc(x.ua||"")}</td></tr>`).join("")||`<tr><td colspan="4" class="muted">Nenhum registro ainda.</td></tr>`}
    </tbody></table></div>`:`<div><button type="button" class="btn primary" id="accLoad">Carregar histórico</button></div>`}</div>`;
}
async function loadAccLog(){ const ids=S.pessoas.map(p=>p.id); const out=[];
  for(const u of ids){ try{ const q=await db.doc("log/"+u).collection("meses").get(); q.docs.forEach(d=>(d.data().eventos||[]).forEach(ev=>out.push({...ev,uid:u}))); }catch(e){ console.error(e); } }
  ACC_LOG=out.sort((a,b)=>String(b.t).localeCompare(String(a.t))); refresh(true); }
async function resolveAccNames(){
  const els=[...document.querySelectorAll(".accname,.accopt")]; if(!els.length) return;
  els.forEach(e=>{ const u=e.dataset.uid; const p=S.pessoas.find(x=>x.id===u); e.textContent=u===S.myId?"Você":(p&&(p.nome||p.email))||u; });
}
async function saveAcesso(uid,role){
  await guard(()=> role==="pendente" ? db.doc("niveis/"+uid).delete() : db.doc("niveis/"+uid).set({nivel:role,em:new Date().toISOString(),por:S.myId}),"Acesso atualizado");
}

/* ---------- Cadastros ---------- */
function shellCadastros(){ return `<div id="data"></div>`; }
const LISTS = [["analistas","Analistas / responsáveis","Quem cadastra, retira ou descarta."],["locais","Localizações","Onde as amostras ficam guardadas."],["tipos","Tipos de amostra",""],["motivos","Motivos de retirada",""],["destinos","Destinos",""],["unidades","Unidades",""]];
function dataCadastros(){
  const ro=S.readonly||!FULL();
  if(FULL()&&S.mode==="db") setTimeout(resolveAccNames,0);
  return `${S.isOwner&&S.mode==="db"?`<div class="panel"><h2>Backup <small>só a proprietária</small></h2>
    <p class="sub" style="margin:0"><b>Importar do app do Claude:</b> escolha o arquivo .json baixado em Cadastros → Backup completo no app antigo. Pode repetir sem duplicar: registros iguais são substituídos.</p>
    <div class="toolbar"><div class="field grow"><label for="bkfile">Arquivo de backup (.json)</label><input id="bkfile" type="file" accept=".json,application/json"></div></div><div id="bkprev">${BKMSG?`<p style="margin:0">${BKMSG}</p>`:""}</div>
    <p class="sub" style="margin:0"><b>Cópia de segurança:</b> baixa todos os dados deste app em um arquivo .json.</p><div><button class="btn" data-act="backup">Baixar backup (.json)</button></div></div>`:""}${S.isOwner&&S.mode==="db"?acessosPanel():""}${FULL()&&S.mode==="db"?historicoAcessoPanel():""}<div class="panel"><h2>Importar planilha <small>.xlsx ou .csv</small></h2>
    <p style="margin:0" class="muted">A primeira aba é lida. As colunas são reconhecidas pelo nome do cabeçalho: Produto (ou Material), Registro MAPA, Código interno, Lote, Quantidade, Data de Fabricação, Data de Validade, Localização, Tipo, Unidade, Responsável, Observação. Cada linha vira um lote com uma entrada inicial igual à quantidade.</p>
    <div class="toolbar"><div class="field grow"><label for="impfile">Arquivo</label><input id="impfile" type="file" accept=".xlsx,.xls,.csv"${ro?" disabled":""}></div></div>
    <div id="impprev"></div></div>
  <div class="panel"><h2>Alerta de vencimento</h2>
    <div class="toolbar"><div class="field" style="flex:0 1 220px"><label for="cfgal">Avisar com quantos dias de antecedência</label><input id="cfgal" type="number" min="1" max="730" value="${esc(S.cfg.alerta)}"${ro?" disabled":""}></div>
    <button class="btn" data-act="savealerta"${ro?" disabled":""}>Salvar</button></div>
    <p class="sub" style="margin:0">Mesma regra da planilha: acima do prazo fica “Ok”, dentro do prazo “A vencer”, na data ou depois “Vencida”.</p></div>
  <div class="panel"><h2>Produtos acabados <small>${PRODS().length} cadastrados</small></h2>
    <p class="sub" style="margin:0">Ao escolher o produto no cadastro da amostra, o registro MAPA é preenchido sozinho.</p>
    ${ro?"":`<form class="addprod toolbar"><div class="field grow"><label for="np_nome">Nome comercial</label><input id="np_nome" required></div><div class="field"><label for="np_mapa">Registro MAPA</label><input id="np_mapa" placeholder="CE 001014-6.000000" style="font-family:var(--mono)"></div><div class="field"><label for="np_cat">Categoria</label><select id="np_cat"><option>Orgânico</option><option>Convencional</option></select></div><button class="btn" type="submit">Adicionar produto</button></form>`}
    <div class="tablebox"><table><thead><tr><th>Produto</th><th>Registro MAPA</th><th>Categoria</th><th>Par para reclassificação</th><th></th></tr></thead><tbody>
    ${PRODS().map((p,i)=>`<tr class="row"><td class="prod">${esc(p.nome)}</td><td class="mono">${esc(p.mapa||"—")}</td><td>${esc(p.categoria||"")}</td><td>${p.categoria==="Orgânico"&&!ro?`<select class="parsel" data-i="${i}" aria-label="Par convencional de ${esc(p.nome)}" style="max-width:280px"><option value="">—</option>${PRODS().filter(x=>x.categoria==="Convencional").map(x=>`<option${x.nome===p.par?" selected":""}>${esc(x.nome)}</option>`).join("")}</select>`:`<span class="sub">${esc(p.par||"")}</span>`}</td><td><div class="rowacts">${ro?"":`<button class="btn sm" data-act="rmprod" data-i="${i}" aria-label="Remover ${esc(p.nome)}">Remover</button>`}</div></td></tr>`).join("")||`<tr><td colspan="5" class="muted">Nenhum produto cadastrado.</td></tr>`}
    </tbody></table></div></div>
  <div class="lists">${LISTS.map(([k,t,d])=>`<div class="panel"><h2>${t} <small>${(S.cfg[k]||[]).length}</small></h2>${d?`<p class="sub" style="margin:0">${d}</p>`:""}
    <div class="chips">${(S.cfg[k]||[]).map((v,i)=>`<span class="chip">${esc(v)}${ro?"":`<button data-act="rmlist" data-k="${k}" data-i="${i}" aria-label="Remover ${esc(v)}">×</button>`}</span>`).join("")||`<span class="sub">Lista vazia.</span>`}</div>
    ${ro?"":`<form class="addrow" data-k="${k}"><input id="add-${k}" placeholder="Adicionar…" aria-label="Adicionar em ${t}"><button class="btn sm" type="submit">Adicionar</button></form>`}</div>`).join("")}</div>`;
}

/* ======================= Eventos ======================= */
function bindShell(){
  const on=(id,k,obj,ev="input")=>{ const el=$("#"+id); if(el) el.addEventListener(ev,()=>{ obj[k]=el.value; document.querySelectorAll(`.ptab[data-k="${k}"]`).forEach(x=>{ if(S[x.dataset.o]===obj) x.setAttribute("aria-pressed",String(x.dataset.v===el.value)); }); refresh(true); }); };
  on("fq","q",S.f); on("fst","status",S.f,"change"); on("floc","local",S.f,"change"); on("ftp","tipo",S.f,"change"); on("fso","sort",S.f,"change");
  on("cq","q",S.fc); on("cpa","pais",S.fc,"change"); on("ctp","tipo",S.fc,"change");
  on("eq","q",S.fe); on("ecat","cat",S.fe,"change"); on("efo","foto",S.fe,"change"); on("ede","de",S.fe,"change"); on("eat","ate",S.fe,"change");
  on("mq","q",S.fm); on("mtp","tipo",S.fm,"change"); on("mrs","resp",S.fm,"change"); on("mde","de",S.fm,"change"); on("mat","ate",S.fm,"change");
}
$("#main").addEventListener("click", async e=>{
  if(e.target.id==="accLoad"){ e.target.disabled=true; e.target.textContent="Carregando…"; await loadAccLog(); return; }
  const rs=e.target.closest(".accreset"); if(rs){ const u=rs.dataset.uid; confirmar("Enviar e-mail de nova senha",`${u} recebe um link para criar uma senha nova. A senha atual continua valendo até ela trocar.`,"Enviar",async()=>{ await guard(async()=>{ await firebase.auth().sendPasswordResetEmail(u); await logEv("reset_enviado",null,{alvo:u}); },"E-mail enviado"); }); return; }
  const b=e.target.closest("[data-act]"); if(!b) return;
  const id=b.dataset.id, act=b.dataset.act;
  if(act==="pt"){ const o=S[b.dataset.o]; o[b.dataset.k]=b.dataset.v; const map={f:{status:"fst"},fm:{tipo:"mtp"},fc:{tipo:"ctp"},fe:{cat:"ecat"}}; const sel=$("#"+((map[b.dataset.o]||{})[b.dataset.k]||"_")); if(sel) sel.value=b.dataset.v; b.parentElement.querySelectorAll(".ptab").forEach(x=>x.setAttribute("aria-pressed",String(x===b))); MOBLIM.estoque=40; MOBLIM.movs=40; MOBLIM.evid=40; refresh(true); return; }
  if(act==="mais"){ MOBLIM[b.dataset.k]+=40; refresh(true); return; }
  if(act==="qnova"){ $("#btnNova").click(); return; }
  if(act==="qret"){ $("#btnRetirar").click(); return; }
  if(act==="nova") formAmostra();
  else if(act==="gotab") setTab(b.dataset.tab);
  else if(act==="backup") backupJSON(b);
  else if(act==="instalar") instalarApp();
  else if(act==="evid") evidDetalhe(id);
  else if(act==="retirar") formMov(id,"saida");
  else if(act==="entrada") formMov(id,"entrada");
  else if(act==="descartar") formMov(id,"descarte");
  else if(act==="hist") historico(id);
  else if(act==="ajustar") ajustarLote(id);
  else if(act==="novocli") formCliente();
  else if(act==="cli") clienteDetalhe(id);
  else if(act==="enviarcli") formMov(null,"saida",id);
  else if(act==="delmov") confirmar("Desfazer lançamento","Esta movimentação será apagada e o saldo do lote recalculado. As fotos de evidência dela também são apagadas.","Desfazer", ()=>guard(()=>store.deleteMov(id),"Lançamento desfeito"));
  else if(act==="rmlist"){ const k=b.dataset.k,i=+b.dataset.i; const cfg={...S.cfg,[k]:S.cfg[k].filter((_,j)=>j!==i)}; await guard(()=>store.saveCfg(cfg)); refresh(true); }
  else if(act==="rmprod"){ const i=+b.dataset.i; const cfg={...S.cfg,produtos:PRODS().filter((_,j)=>j!==i)}; await guard(()=>store.saveCfg(cfg)); refresh(true); }
  else if(act==="savealerta"){ const v=Math.max(1,parseInt($("#cfgal").value)||60); await guard(()=>store.saveCfg({...S.cfg,alerta:v}),"Alerta salvo"); refresh(true); }
  else if(act==="doimport") doImport();
});
$("#main").addEventListener("submit", async e=>{
  const fp=e.target.closest("form.addprod");
  if(fp){ e.preventDefault(); const nome=$("#np_nome").value.trim(), mapa=$("#np_mapa").value.trim(), categoria=$("#np_cat").value; if(!nome) return;
    if(PRODS().some(x=>norm(x.nome)===norm(nome))){ toast("Esse produto já está cadastrado."); return; }
    const cfg={...S.cfg,produtos:[...PRODS(),{nome,mapa,categoria}].sort((a,b)=>a.nome.localeCompare(b.nome,"pt"))};
    document.activeElement&&document.activeElement.blur(); await guard(()=>store.saveCfg(cfg),"Produto adicionado"); refresh(true); return; }
  const f=e.target.closest("form.addrow"); if(!f) return; e.preventDefault();
  const k=f.dataset.k, inp=f.querySelector("input"); const v=inp.value.trim(); if(!v) return;
  if((S.cfg[k]||[]).some(x=>norm(x)===norm(v))){ toast("Já existe na lista."); return; }
  const cfg={...S.cfg,[k]:[...(S.cfg[k]||[]),v].sort((a,b)=>a.localeCompare(b,"pt"))};
  inp.value=""; inp.blur(); await guard(()=>store.saveCfg(cfg)); refresh(true);
});
$("#main").addEventListener("change", async e=>{ if(e.target.id==="impfile") lerImport(e.target.files[0]);
  if(e.target.id==="bkfile") importarBackup(e.target.files[0]);
  if(e.target.classList&&e.target.classList.contains("accsel")) await saveAcesso(e.target.dataset.uid,e.target.value);
  if(e.target.id==="accF"){ ACC_F=e.target.value; refresh(true); }
  if(e.target.classList&&e.target.classList.contains("parsel")){ const i=+e.target.dataset.i; const L=PRODS().map(x=>({...x})); if(e.target.value) L[i].par=e.target.value; else delete L[i].par; await guard(()=>store.saveCfg({...S.cfg,produtos:L}),"Par salvo"); } });
$("#btnNova").onclick=()=>formAmostra();
$("#btnRetirar").onclick=()=>formMov(null,"saida");
$("#btnExport").onclick=()=>exportar();

/* ======================= Modais ======================= */
function modal(title, body, foot, wide){
  const ov=document.createElement("div"); ov.className="overlay";
  ov.innerHTML=`<div class="modal${wide?" wide":""}" role="dialog" aria-modal="true" aria-label="${esc(title)}"><header><h3>${esc(title)}</h3><button class="x" aria-label="Fechar">×</button></header><form novalidate><div class="body">${body}</div><footer>${foot}</footer></form></div>`;
  document.body.appendChild(ov);
  const close=()=>{ ov.remove(); document.removeEventListener("keydown",k); };
  const k=e=>{ if(e.key==="Escape") close(); };
  document.addEventListener("keydown",k);
  ov.querySelector(".x").onclick=close;
  ov.addEventListener("mousedown",e=>{ if(e.target===ov) close(); });
  ov.querySelectorAll("[data-close]").forEach(b=>b.onclick=close);
  setTimeout(()=>{ const f=ov.querySelector("input:not([type=hidden]):not([readonly]),select"); f&&f.focus(); },30);
  return {ov, form:ov.querySelector("form"), close};
}
function confirmar(title,msg,label,fn){
  const m=modal(title,`<p class="full" style="margin:0">${esc(msg)}</p>`,`<button type="button" class="btn" data-close>Cancelar</button><button type="submit" class="btn danger">${esc(label)}</button>`);
  m.form.onsubmit=async e=>{ e.preventDefault(); m.close(); await fn(); };
}
const dl=(id,list)=>`<datalist id="${id}">${[...new Set(list)].filter(Boolean).map(v=>`<option value="${esc(v)}">`).join("")}</datalist>`;
const lastResp=()=>LS.get("resp","");

function formAmostra(a){
  if(S.readonly) return; if(a&&!FULL()) return;
  const ed=!!a; a=a||{tipo:S.cfg.tipos[0],unidade:S.cfg.unidades[0],responsavel:lastResp(),fabricacao:hoje()};
  const prods=[...PRODS().map(p=>p.nome),...S.amostras.map(x=>x.produto)];
  const locs=[...S.cfg.locais,...S.amostras.map(x=>x.local)];
  const body=`
   <div class="field full"><label for="a_prod">Produto *</label><input id="a_prod" list="dlp" required value="${esc(a.produto)}" placeholder="Nome do produto acabado">${dl("dlp",prods)}</div>
   <div class="field"><label for="a_mapa">Registro MAPA</label><input id="a_mapa" value="${esc(a.mapa)}" style="font-family:var(--mono)"></div>
   <div class="field"><label for="a_cod">Código interno</label><input id="a_cod" value="${esc(a.codigo)}"></div>
   <div class="field"><label for="a_lote">Lote *</label><input id="a_lote" required value="${esc(a.lote)}" style="font-family:var(--mono)"></div>
   <div class="field"><label for="a_tipo">Tipo de amostra</label><select id="a_tipo">${opts(S.cfg.tipos,a.tipo)}</select></div>
   <div class="field"><label for="a_cont">Conteúdo por unidade</label><input id="a_cont" value="${esc(a.conteudo)}" placeholder="ex.: 1 kg, 500 mL"></div>
   <div class="field"><label for="a_fab">Data de fabricação</label><input id="a_fab" type="date" value="${esc(a.fabricacao)}"></div>
   <div class="field"><label for="a_val">Data de validade *</label><input id="a_val" type="date" required value="${esc(a.validade)}"></div>
   ${ed?"":`<div class="field"><label for="a_qtd">Quantidade guardada *</label><input id="a_qtd" type="number" min="0" step="any" inputmode="decimal" required></div>`}
   <div class="field"><label for="a_un">Unidade</label><select id="a_un">${opts(S.cfg.unidades,a.unidade)}</select></div>
   <div class="field"><label for="a_loc">Localização</label><input id="a_loc" list="dll" value="${esc(a.local)}">${dl("dll",locs)}</div>
   <div class="field"><label for="a_resp">Responsável</label><select id="a_resp">${opts(S.cfg.analistas,a.responsavel,"—")}</select></div>
   ${ed?"":`<div class="field full"><label for="a_ids">ID das amostras</label><input id="a_ids" placeholder="ex.: A1, A2, B1, B2" style="font-family:var(--mono)"></div><div class="field"><label for="a_kgun">Peso por amostra (kg)</label><input id="a_kgun" type="number" step="0.001" min="0" inputmode="decimal" placeholder="ex.: 0,345"></div><div class="field"></div>`}
   <div class="field full"><label for="a_obs">Observação</label><textarea id="a_obs">${esc(a.obs)}</textarea></div>
   <div class="formerr full" id="a_err"></div>`;
  const foot=`${ed?`<button type="button" class="btn danger" id="a_del" style="margin-right:auto">Excluir cadastro</button>`:""}<button type="button" class="btn" data-close>Cancelar</button><button type="submit" class="btn primary">${ed?"Salvar alterações":"Cadastrar amostra"}</button>`;
  const m=modal(ed?"Editar amostra":"Nova amostra",body,foot);
  m.ov.querySelector("#a_prod").addEventListener("input",ev=>{ const mp=mapaDe(ev.target.value); if(mp) m.ov.querySelector("#a_mapa").value=mp; });
  if(ed) m.ov.querySelector("#a_del").onclick=()=>{ m.close(); confirmar("Excluir cadastro",`Apaga o lote ${a.lote} de ${a.produto} e todas as suas movimentações. Use só para cadastro feito por engano; para amostra usada ou vencida, registre a retirada ou o descarte.`,"Excluir",()=>guard(()=>store.deleteAmostra(a.id),"Cadastro excluído")); };
  m.form.onsubmit=async e=>{
    e.preventDefault(); const g=id=>m.ov.querySelector("#"+id).value.trim();
    const d={produto:g("a_prod"),mapa:g("a_mapa"),codigo:g("a_cod"),lote:g("a_lote"),tipo:g("a_tipo"),conteudo:g("a_cont"),fabricacao:g("a_fab"),validade:g("a_val"),unidade:g("a_un"),local:g("a_loc"),responsavel:g("a_resp"),obs:g("a_obs")};
    const err=m.ov.querySelector("#a_err");
    if(!d.produto||!d.lote||!d.validade){ err.textContent="Preencha produto, lote e data de validade."; return; }
    if(d.fabricacao&&d.validade<d.fabricacao){ err.textContent="A validade não pode ser anterior à fabricação."; return; }
    let qtd=0; if(!ed){ qtd=Number(g("a_qtd").replace(",",".")); if(!(qtd>0)){ err.textContent="Informe a quantidade guardada (maior que zero)."; return; } }
    const dup=S.amostras.find(x=>x.id!==a.id&&norm(x.produto)===norm(d.produto)&&norm(x.lote)===norm(d.lote)&&x.tipo===d.tipo&&!x.descartada);
    if(dup&&!err.dataset.dupok){ err.textContent="Já existe esse produto/lote/tipo em estoque. Para somar unidades use “+ Entrada” no lote existente. Clique de novo para cadastrar mesmo assim."; err.dataset.dupok="1"; return; }
    if(d.responsavel) LS.set("resp",d.responsavel);
    m.close();
    if(ed) await guard(()=>store.updateAmostra(a.id,{...d,atualizadoEm:hoje()}),"Amostra atualizada");
    else { const ni=parseIds(g("a_ids")); if(ni.length){ d.ids=ni; const ku=Number(g("a_kgun"))||0; if(ku){ d.idsKg={}; ni.forEach(i=>d.idsKg[i]=ku); } } await guard(()=>store.addAmostra(d,qtd,d.responsavel),"Amostra cadastrada"); }
  };
}

function formMov(id, tipo, clienteId){
  if(S.readonly) return;
  const R=rows(); let r=id?R.find(x=>x.id===id):null;
  const elig=R.filter(x=>tipo==="entrada"?!x.descartada:(x.saldo>0&&!x.descartada)).sort((a,b)=>a.produto.localeCompare(b.produto,"pt")||String(a.lote).localeCompare(String(b.lote)));
  if(!r&&!elig.length){ toast("Não há amostras com saldo para retirar."); return; }
  const titles={saida:"Registrar retirada",entrada:"Registrar entrada",descarte:"Registrar descarte"};
  const motivos= tipo==="descarte"? S.cfg.motivos.filter(m=>/descarte/i.test(m)) : tipo==="entrada"? ["Complemento de amostra","Devolução","Ajuste de inventário"] : S.cfg.motivos.filter(m=>!/descarte/i.test(m));
  const sel = r? `<div class="ref full" id="m_ref"></div><input type="hidden" id="m_am" value="${r.id}">`
    : `<div class="field full"><label for="m_am">Amostra (produto · lote)</label><select id="m_am">${elig.map(x=>`<option value="${x.id}">${esc(x.produto)} · lote ${esc(x.lote)} · saldo ${nf(x.saldo)} ${esc(x.unidade||"")}${x.st==="vencida"?" (vencida)":""}</option>`).join("")}</select></div><div class="ref full" id="m_ref"></div>`;
  const out=tipo!=="entrada";
  const catIni=clienteId?(isLab(clienteId)?"Laboratório externo":"Cliente"):"";
  const body=`${sel}
    ${tipo==="saida"?`<div class="field full"><label for="m_cat">Tipo de saída *</label><select id="m_cat"><option value="">Selecione</option>${CATS_SAIDA.map(c=>`<option${c===catIni?" selected":""}>${c}</option>`).join("")}</select></div>`:""}
    <div class="field"><label for="m_qtd">Quantidade *</label><input id="m_qtd" type="number" min="0" step="any" inputmode="decimal" required></div>
    <div class="field"><label for="m_data">Data *</label><input id="m_data" type="date" value="${hoje()}" required></div>
    <div class="field"><label for="m_mot">Motivo</label><select id="m_mot">${opts(motivos.length?motivos:["Outro"],clienteId?(motivos.find(x=>(isLab(clienteId)?/externa/i:/cliente/i).test(x))||""):"")}</select></div>
    ${tipo==="saida"?`<div class="field"><label for="m_dest">Destino</label><select id="m_dest">${opts(S.cfg.destinos,clienteId?(S.cfg.destinos.find(x=>/extern/i.test(x))||""):"","—")}</select></div>
    <div class="field full" id="m_cliwrap"><label for="m_cli">Cliente / laboratório</label><select id="m_cli"><option value="">—</option>${S.clientes.slice().sort((a,b)=>a.nome.localeCompare(b.nome,"pt")).map(c=>`<option value="${c.id}"${c.id===clienteId?" selected":""}>${esc(c.nome)}${c.pais?" · "+esc(c.pais):""}</option>`).join("")}</select></div>`:`<div class="field"></div>`}
    <div class="field full"><label for="m_ids">ID da(s) amostra(s)</label><div class="chips" id="m_idchips" style="margin-bottom:6px"></div>
      ${out?`<div id="m_selbox" style="display:flex;flex-direction:column;gap:6px;margin-bottom:6px"></div>`:""}
      <input id="m_ids" placeholder="${out?"Clique nos IDs acima ou digite, ex.: A125x1, B124":"IDs das novas amostras, ex.: A120, A121"}" style="font-family:var(--mono)"><div class="sub" id="m_kgsel" style="margin-top:4px"></div></div>
    ${tipo==="entrada"?`<div class="field"><label for="m_kgun">Peso por amostra (kg)</label><input id="m_kgun" type="number" step="0.001" min="0" inputmode="decimal" placeholder="ex.: 0,345"></div><div class="field"></div>`:`<div class="field"><label for="m_kg">Peso retirado (kg)</label><input id="m_kg" inputmode="decimal" autocomplete="off" placeholder="Cole ou digite, ex.: 0,250" style="font-family:var(--mono)"></div><div class="field" style="justify-content:flex-end"><div class="sub" id="m_kgparse">Aceita vírgula ou ponto. Vários pesos colados (um por linha) são somados.</div></div>`}
    <div class="field full"><label for="m_resp">Responsável *</label><select id="m_resp">${opts(S.cfg.analistas,lastResp(),"Selecione")}</select></div>
    <div class="field full"><label for="m_obs" id="m_obsl">Observação</label><input id="m_obs" placeholder="${tipo==="saida"?"ex.: nº da análise, cliente, protocolo de reclamação":""}"></div>
    ${out?fotoBox(tipo==="descarte"?[["saida","Fotos do descarte"]]:[["saida","Fotos da saída"],["preparo","Fotos do preparo"]]):""}
    <div class="formerr full" id="m_err"></div>`;
  const m=modal(titles[tipo],body,`<button type="button" class="btn" data-close>Cancelar</button><button type="submit" class="btn ${tipo==="descarte"?"danger":"primary"}">${titles[tipo].replace("Registrar ","Confirmar ")}</button>`);
  const q=s=>m.ov.querySelector(s);
  const fotos=out?bindFotoBox(m.ov):[];
  const catSync=()=>{ const c=q("#m_cat"); if(!c) return; const outra=c.value==="Outras saídas"; q("#m_obsl").textContent=outra?"Para onde foram as amostras? *":"Observação"; q("#m_obs").placeholder=outra?"ex.: feira, visita técnica, setor comercial":"ex.: nº da análise, cliente, protocolo de reclamação"; };
  if(q("#m_cat")){ q("#m_cat").addEventListener("change",catSync); catSync(); }
  /* seleção por ID: quantos volumes saem de cada ID */
  let selQ={};
  const unOf=(io,i)=>(io.q&&io.q[i])||1;
  const kgUnit=(io,i)=>(Number(io.kg[i])||0)/unOf(io,i);
  const tokens=()=>{ const io=idsOf(r); return Object.keys(selQ).map(i=>unOf(io,i)>1?`${i}x${selQ[i]}`:i).join(", "); };
  const parseSel=t=>{ const io=idsOf(r), o={}; String(t||"").split(/[,;\s]+/).map(x=>x.trim().toUpperCase()).filter(Boolean).forEach(x=>{ const mm=x.match(/^(.+?)(?:[X×*](\d+))?$/); if(!mm) return; const i=mm[1]; const n=mm[2]?parseInt(mm[2]):1; if(!io.ids.includes(i)) return; o[i]=Math.max(1,Math.min(unOf(io,i),n)); }); return o; };
  const estKg=()=>{ const io=idsOf(r); return Object.keys(selQ).reduce((s,i)=>s+kgUnit(io,i)*selQ[i],0); };
  const selUnits=()=>Object.values(selQ).reduce((s,n)=>s+n,0);
  const paint=(fromInput)=>{
    const io=idsOf(r);
    if(out){
      if(!fromInput) q("#m_ids").value=tokens();
      const box=q("#m_selbox");
      box.innerHTML=Object.keys(selQ).map(i=>{ const un=unOf(io,i), n=selQ[i], ku=kgUnit(io,i);
        return `<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:var(--sunk);border-radius:6px;padding:6px 10px">
          <b style="font-family:var(--mono);min-width:56px">${esc(i)}</b>
          ${un>1?`<button type="button" class="btn sm selm" data-id="${esc(i)}" aria-label="Menos um volume">−</button><span style="font-family:var(--mono);min-width:20px;text-align:center"><b>${n}</b></span><button type="button" class="btn sm selp" data-id="${esc(i)}" aria-label="Mais um volume">+</button><span class="sub">de ${un} volume${un===1?"":"s"}</span>`:`<span class="sub">1 volume</span>`}
          <span class="sub" style="margin-left:auto">${ku?`${fkg(ku)} kg por volume · sai ${fkg(ku*n)} kg`:"sem peso registrado"}${un>1&&n<un?` · ficam ${un-n} (${fkg(ku*(un-n))} kg)`:""}</span>
          <button type="button" class="btn sm selx" data-id="${esc(i)}" aria-label="Tirar ${esc(i)} da seleção">×</button></div>`; }).join("");
      const kk=estKg(); const kd=q("#m_kgsel"); kd.textContent=selUnits()?`${selUnits()} volume${selUnits()===1?"":"s"} selecionado${selUnits()===1?"":"s"} · peso estimado ${fkg(kk)} kg`:"";
      if(selUnits()) q("#m_qtd").value=selUnits();
      const kin=q("#m_kg"); if(kin&&!kin.dataset.touched){ kin.value=kk?String(Math.round(kk*1000)/1000).replace(".",","):""; showKg(); }
      m.ov.querySelectorAll(".idc").forEach(b=>{ const on=!!selQ[b.dataset.id]; b.style.background=on?"var(--accent)":"var(--sunk)"; b.style.color=on?"var(--accent-ink)":"var(--ink)"; });
    }
  };
  const upd=()=>{ r=R.find(x=>x.id===q("#m_am").value); selQ={}; q("#m_ref").innerHTML=`<div><div class="prod">${esc(r.produto)}</div><div class="sub">Lote <span class="lote">${esc(r.lote)}</span> · validade ${fmtD(r.validade)} · ${esc(r.local||"sem local")}</div></div><div>${pill(r.st)} <b style="font-family:var(--mono)">${nf(r.saldo)}</b> <span class="sub">${esc(r.unidade||"")} em estoque</span></div>`;
    if(tipo==="descarte") q("#m_qtd").value=r.saldo;
    const io=idsOf(r), box=q("#m_idchips");
    if(!out||!io.ids.length){ box.innerHTML=out?`<span class="sub">Nenhum ID registrado para este lote.</span>`:""; }
    else { box.innerHTML=`<span class="sub" style="width:100%">IDs em estoque (${io.ids.length}) · clique para retirar 1 volume${io.conf?` · <span style="color:var(--warn)">conferir: ${esc(io.conf)}</span>`:""}</span>`+io.ids.map(i=>`<button type="button" class="chip idc" data-id="${esc(i)}" style="padding:3px 10px;cursor:pointer;font-family:var(--mono);border:1px solid var(--line)">${esc(idLabel(i,io.q,io.kg))}</button>`).join("");
      if(tipo==="descarte"){ io.ids.forEach(i=>selQ[i]=unOf(io,i)); } }
    paint(); };
  m.ov.addEventListener("click",ev=>{ const io=idsOf(r);
    const c=ev.target.closest(".idc"); if(c){ const i=c.dataset.id; if(selQ[i]) delete selQ[i]; else selQ[i]=1; paint(); return; }
    const p=ev.target.closest(".selp"), mi=ev.target.closest(".selm"), x=ev.target.closest(".selx");
    if(p){ const i=p.dataset.id; selQ[i]=Math.min(unOf(io,i),selQ[i]+1); paint(); }
    else if(mi){ const i=mi.dataset.id; selQ[i]=Math.max(1,selQ[i]-1); paint(); }
    else if(x){ delete selQ[x.dataset.id]; paint(); } });
  const showKg=()=>{ const kin=q("#m_kg"), o=q("#m_kgparse"); if(!kin||!o) return; const v=parseKg(kin.value);
    o.innerHTML=!kin.value.trim()?"Aceita vírgula ou ponto. Vários pesos colados (um por linha) são somados.":v==null?`<span style="color:var(--bad)">Não entendi esse peso. Use números, ex.: 0,250</span>`:`Peso a registrar: <b style="font-family:var(--mono)">${fkg(v)} kg</b>`; };
  m.ov.addEventListener("input",ev=>{ if(ev.target.id==="m_ids"&&out){ selQ=parseSel(ev.target.value); paint(true); } if(ev.target.id==="m_kg"){ ev.target.dataset.touched=ev.target.value.trim()?"1":""; showKg(); } });
  m.ov.addEventListener("paste",ev=>{ if(ev.target.id!=="m_kg") return; const t=(ev.clipboardData||window.clipboardData)?.getData("text"); if(!t) return; ev.preventDefault();
    const v=parseKg(t); ev.target.value=v==null?t.trim():String(v).replace(".",","); ev.target.dataset.touched="1"; showKg(); });
  upd(); if(!id) q("#m_am").onchange=upd;
  m.form.onsubmit=async e=>{
    e.preventDefault(); const err=q("#m_err");
    const qtd=Number(q("#m_qtd").value.replace(",",".")); const resp=q("#m_resp").value; const data=q("#m_data").value;
    if(!(qtd>0)){ err.textContent="Informe uma quantidade maior que zero."; return; }
    if(out&&qtd>r.saldo+1e-9){ err.textContent=`A quantidade passa do saldo do lote (${nf(r.saldo)}).`; return; }
    if(!resp){ err.textContent="Selecione o responsável. Se o nome não está na lista, cadastre em Cadastros → Analistas."; return; }
    if(!data){ err.textContent="Informe a data."; return; }
    const kgTxt=q("#m_kg")?q("#m_kg").value.trim():""; const kgVal=kgTxt?parseKg(kgTxt):null;
    if(kgTxt&&kgVal==null){ err.textContent="O peso retirado não é um número válido. Use, por exemplo, 0,250."; return; }
    if(out&&selUnits()&&Math.abs(qtd-selUnits())>1e-9){ err.textContent=`Os volumes marcados somam ${selUnits()}, mas a quantidade é ${nf(qtd)}. Ajuste os volumes com − e + ou corrija a quantidade.`; return; }
    const cliId=q("#m_cli")?q("#m_cli").value:""; const cli=S.clientes.find(c=>c.id===cliId);
    const cat=tipo==="descarte"?"Descarte":tipo==="saida"?q("#m_cat").value:"";
    if(tipo==="saida"&&!cat){ err.textContent="Escolha o tipo de saída (análises, degustação, laboratório externo, carregamento, cliente ou outras)."; return; }
    if(cat==="Outras saídas"&&!q("#m_obs").value.trim()){ err.textContent="Em “Outras saídas”, informe para onde foram as amostras."; q("#m_obs").focus(); return; }
    if((cat==="Cliente"||cat==="Laboratório externo")&&!cli){ err.textContent=`Escolha ${cat==="Cliente"?"o cliente":"o laboratório"} que vai receber a amostra. Se não está na lista, cadastre na aba Clientes / Lab. externo.`; return; }
    if(tipo==="saida"&&/cliente|externa/i.test(q("#m_mot").value)&&!cli){ err.textContent="Escolha o cliente ou laboratório que vai receber a amostra. Se não está na lista, cadastre na aba Clientes / Lab. externo."; return; }
    LS.set("resp",resp); m.close();
    const extra=cli?{clienteId:cli.id,cliente:cli.nome}:{};
    const io=idsOf(r);
    const est=estKg();
    const kgMov= out ? Math.round(((kgVal!=null)?kgVal:est)*1000)/1000 : (()=>{ const ku=Number((q("#m_kgun")||{}).value)||0; return ku?Math.round(ku*parseIds(q("#m_ids").value).length*1000)/1000:0; })();
    const idsTxt= out ? tokens() : q("#m_ids").value.trim();
    const movDoc={amostraId:r.id,produto:r.produto,lote:r.lote,tipo,qtd,data,motivo:q("#m_mot").value,destino:q("#m_dest")?q("#m_dest").value:"",responsavel:resp,obs:q("#m_obs").value.trim(),ids:idsTxt,kg:kgMov,...(cat?{categoria:cat}:{}),...extra};
    let movId=null;
    const ok=await guard(async()=>{ movId=await store.addMov(movDoc); },
      tipo==="saida"?"Retirada registrada":tipo==="entrada"?"Entrada registrada":"Descarte registrado");
    if(ok&&movId&&fotos.length){ const nf0=fotos.length; toast(`Salvando ${nf0} foto${nf0===1?"":"s"}…`); salvarFotos({id:movId,...movDoc},fotos.splice(0)).then(n=>toast(n===nf0?`${n} foto${n===1?"":"s"} de evidência salva${n===1?"":"s"}`:`Só ${n} de ${nf0} fotos foram salvas. Tente de novo em Evidências.`)); }
    if(ok){
      let ids=[...io.ids], qq={...io.q}, kg={...io.kg}, changed=false;
      if(!out){ const typed=parseIds(q("#m_ids").value); const ku=Number((q("#m_kgun")||{}).value)||0; for(const i of typed){ if(!ids.includes(i)) ids.push(i); if(ku) kg[i]=ku; changed=true; } }
      else for(const i of Object.keys(selQ)){
        const un=unOf(io,i), n=selQ[i], k0=Number(kg[i])||0;
        const share= kgVal!=null ? (est>0 ? kgVal*(kgUnit(io,i)*n)/est : kgVal/Object.keys(selQ).length) : kgUnit(io,i)*n;
        if(n>=un){ ids=ids.filter(x=>x!==i); delete qq[i]; delete kg[i]; }
        else { const rest=un-n; if(rest>1) qq[i]=rest; else delete qq[i]; kg[i]=Math.max(0,Math.round((k0-share)*1000)/1000); }
        changed=true; }
      if(changed) await guard(()=>store.updateAmostra(r.id,{ids,idsQ:qq,idsKg:kg,idsConf:io.conf||"",atualizadoEm:data}));
    }
    if(ok && tipo==="descarte" && Math.abs(qtd-r.saldo)<1e-9) await guard(()=>store.updateAmostra(r.id,{descartada:true,atualizadoEm:data}));
  };
}

/* Reclassificação: sai de um produto (ex.: orgânico) e entra em outro (ex.: convencional), mesmo lote */
function parDe(nome){ const p=PRODS().find(x=>norm(x.nome)===norm(nome)); return p&&p.par?p.par:""; }
function reclassificar(id){
  if(S.readonly||!FULL()) return;
  const R=rows(); const r=R.find(x=>x.id===id); if(!r) return;
  const io=idsOf(r);
  const prods=PRODS().filter(p=>norm(p.nome)!==norm(r.produto));
  const sug=parDe(r.produto);
  const conv=prods.filter(p=>p.categoria==="Convencional"), outros=prods.filter(p=>p.categoria!=="Convencional");
  const optP=list=>list.sort((a,b)=>a.nome.localeCompare(b.nome,"pt")).map(p=>`<option value="${esc(p.nome)}"${p.nome===sug?" selected":""}>${esc(p.nome)}${p.mapa?" · "+esc(p.mapa):""}</option>`).join("");
  const body=`<div class="ref full"><div><div class="sub">Sai de</div><div class="prod">${esc(r.produto)}</div><div class="sub">Lote <span class="lote">${esc(r.lote)}</span>${r.mapa?" · MAPA "+esc(r.mapa):""} · validade ${fmtD(r.validade)}</div></div><div>${pill(r.st)} <b style="font-family:var(--mono)">${nf(r.saldo)}</b> <span class="sub">${esc(r.unidade||"")} · ${fkg(kgOf(r))} kg</span></div></div>
   <div class="field full"><label for="rc_dest">Entra em (produto de destino) *</label><select id="rc_dest"><option value="">Selecione</option>${conv.length?`<optgroup label="Convencional">${optP(conv)}</optgroup>`:""}${outros.length?`<optgroup label="Outros">${optP(outros)}</optgroup>`:""}</select>
     <div class="sub" id="rc_destinfo" style="margin-top:4px"></div></div>
   <div class="field full"><label>Volumes que mudam de produto</label>
     ${io.ids.length?`<div class="chips" id="rc_chips" style="margin-bottom:6px"><span class="sub" style="width:100%">Clique para marcar o ID inteiro · <button type="button" class="btn sm" id="rc_all">Marcar todos</button></span>${io.ids.map(i=>`<button type="button" class="chip rcc" data-id="${esc(i)}" style="padding:3px 10px;cursor:pointer;font-family:var(--mono);border:1px solid var(--line)">${esc(idLabel(i,io.q,io.kg))}</button>`).join("")}</div><div id="rc_sel" style="display:flex;flex-direction:column;gap:6px"></div>`:`<span class="sub">Este lote não tem IDs registrados. Informe só a quantidade.</span>`}</div>
   <div class="field"><label for="rc_qtd">Quantidade *</label><input id="rc_qtd" type="number" min="1" step="1" value="${io.ids.length?"":r.saldo}"></div>
   <div class="field"><label for="rc_kg">Peso (kg)</label><input id="rc_kg" inputmode="decimal" placeholder="calculado pelos IDs" style="font-family:var(--mono)"></div>
   <div class="field"><label for="rc_data">Data *</label><input id="rc_data" type="date" value="${hoje()}"></div>
   <div class="field"><label for="rc_mot">Motivo *</label><select id="rc_mot">${opts(["Solicitação do cliente","Acordo comercial","Perda de certificação orgânica","Outro"])}</select></div>
   <div class="field full"><label for="rc_cli">Cliente / acordo (opcional)</label><select id="rc_cli"><option value="">—</option>${S.clientes.filter(c=>(c.tipo||"Cliente")==="Cliente").sort((a,b)=>a.nome.localeCompare(b.nome,"pt")).map(c=>`<option value="${c.id}">${esc(c.nome)}${c.pais?" · "+esc(c.pais):""}</option>`).join("")}</select></div>
   <div class="field full"><label for="rc_resp">Responsável *</label><select id="rc_resp">${opts(S.cfg.analistas,lastResp(),"Selecione")}</select></div>
   <div class="field full"><label for="rc_obs">Observação</label><input id="rc_obs" placeholder="ex.: nº do pedido, e-mail de autorização"></div>
   <div class="ref full" id="rc_resumo" style="display:block"></div>
   <div class="formerr full" id="rc_err"></div>`;
  const m=modal("Reclassificar produto",body,`<button type="button" class="btn" data-close>Cancelar</button><button type="submit" class="btn primary">Confirmar reclassificação</button>`,true);
  const q=s=>m.ov.querySelector(s);
  let selQ={}; const unOf=i=>(io.q&&io.q[i])||1; const kgU=i=>(Number(io.kg[i])||0)/unOf(i);
  const destLot=()=>{ const d=q("#rc_dest").value; return d?S.amostras.find(a=>norm(a.produto)===norm(d)&&norm(a.lote)===norm(r.lote)&&!a.descartada):null; };
  const paint=()=>{
    const box=q("#rc_sel");
    if(box) box.innerHTML=Object.keys(selQ).map(i=>{ const un=unOf(i), n=selQ[i];
      return `<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:var(--sunk);border-radius:6px;padding:6px 10px"><b style="font-family:var(--mono);min-width:56px">${esc(i)}</b>
        ${un>1?`<button type="button" class="btn sm rcm" data-id="${esc(i)}" aria-label="Menos um volume">−</button><b style="font-family:var(--mono)">${n}</b><button type="button" class="btn sm rcp" data-id="${esc(i)}" aria-label="Mais um volume">+</button><span class="sub">de ${un}</span>`:`<span class="sub">1 volume</span>`}
        <span class="sub" style="margin-left:auto">${fkg(kgU(i)*n)} kg</span><button type="button" class="btn sm rcx" data-id="${esc(i)}" aria-label="Desmarcar ${esc(i)}">×</button></div>`; }).join("");
    m.ov.querySelectorAll(".rcc").forEach(b=>{ const on=!!selQ[b.dataset.id]; b.style.background=on?"var(--accent)":"var(--sunk)"; b.style.color=on?"var(--accent-ink)":"var(--ink)"; });
    const un=Object.values(selQ).reduce((s,n)=>s+n,0), kk=Object.keys(selQ).reduce((s,i)=>s+kgU(i)*selQ[i],0);
    if(io.ids.length){ q("#rc_qtd").value=un||""; const kin=q("#rc_kg"); if(!kin.dataset.touched) kin.value=kk?String(Math.round(kk*1000)/1000).replace(".",","):""; }
    const d=q("#rc_dest").value, dl=destLot(), qt=Number(q("#rc_qtd").value)||0, kgv=parseKg(q("#rc_kg").value);
    q("#rc_destinfo").innerHTML=d?(dl?`Já existe o lote ${esc(r.lote)} de ${esc(d)} (saldo ${nf(rows().find(x=>x.id===dl.id)?.saldo||0)}). Os volumes serão somados a ele.`:`Será criado o lote ${esc(r.lote)} de ${esc(d)}${mapaDe(d)?" · MAPA "+esc(mapaDe(d)):""}, com a mesma fabricação, validade e localização.`)+(sug&&d!==sug?` <span style="color:var(--warn)">O par cadastrado deste produto é ${esc(sug)}.</span>`:""):"";
    q("#rc_resumo").innerHTML=d&&qt?`<b>Resumo:</b> sai ${nf(qt)} ${esc(r.unidade||"un")}${kgv?` (${fkg(kgv)} kg)`:""} de <b>${esc(r.produto)}</b> e entra em <b>${esc(d)}</b>, lote ${esc(r.lote)}. Saldo de origem passa de ${nf(r.saldo)} para ${nf(r.saldo-qt)}.`:`<span class="sub">Escolha o produto de destino e os volumes.</span>`;
  };
  m.ov.addEventListener("click",ev=>{
    const c=ev.target.closest(".rcc"); if(c){ const i=c.dataset.id; if(selQ[i]) delete selQ[i]; else selQ[i]=unOf(i); paint(); return; }
    if(ev.target.id==="rc_all"){ io.ids.forEach(i=>selQ[i]=unOf(i)); paint(); return; }
    const p=ev.target.closest(".rcp"), mi=ev.target.closest(".rcm"), x=ev.target.closest(".rcx");
    if(p){ selQ[p.dataset.id]=Math.min(unOf(p.dataset.id),selQ[p.dataset.id]+1); paint(); }
    else if(mi){ selQ[mi.dataset.id]=Math.max(1,selQ[mi.dataset.id]-1); paint(); }
    else if(x){ delete selQ[x.dataset.id]; paint(); } });
  m.ov.addEventListener("input",ev=>{ if(ev.target.id==="rc_kg") ev.target.dataset.touched=ev.target.value.trim()?"1":""; paint(); });
  m.ov.addEventListener("change",ev=>{ if(ev.target.id==="rc_dest") paint(); });
  m.ov.addEventListener("paste",ev=>{ if(ev.target.id!=="rc_kg") return; const t=(ev.clipboardData||window.clipboardData)?.getData("text"); if(!t) return; ev.preventDefault(); const v=parseKg(t); ev.target.value=v==null?t.trim():String(v).replace(".",","); ev.target.dataset.touched="1"; paint(); });
  paint();
  m.form.onsubmit=async e=>{ e.preventDefault(); const err=q("#rc_err");
    const dest=q("#rc_dest").value, qtd=Number(q("#rc_qtd").value), data=q("#rc_data").value, resp=q("#rc_resp").value, mot=q("#rc_mot").value;
    const kgTxt=q("#rc_kg").value.trim(), kgVal=kgTxt?parseKg(kgTxt):null;
    const selUn=Object.values(selQ).reduce((s,n)=>s+n,0);
    if(!dest){ err.textContent="Escolha o produto de destino."; return; }
    if(!(qtd>0)||Math.round(qtd)!==qtd){ err.textContent="Informe a quantidade (número inteiro maior que zero)."; return; }
    if(qtd>r.saldo){ err.textContent=`A quantidade passa do saldo do lote (${nf(r.saldo)}).`; return; }
    if(io.ids.length&&selUn!==qtd){ err.textContent="Marque os IDs que mudam de produto (a quantidade vem da soma dos volumes marcados)."; return; }
    if(kgTxt&&kgVal==null){ err.textContent="O peso não é um número válido. Use, por exemplo, 6,035."; return; }
    if(!data){ err.textContent="Informe a data."; return; }
    if(!resp){ err.textContent="Selecione o responsável."; return; }
    LS.set("resp",resp); m.close();
    const cli=S.clientes.find(c=>c.id===q("#rc_cli").value); const obs=q("#rc_obs").value.trim();
    const est=Object.keys(selQ).reduce((s,i)=>s+kgU(i)*selQ[i],0); const kgMov=Math.round(((kgVal!=null)?kgVal:est)*1000)/1000;
    // volumes que mudam
    const moved=Object.keys(selQ).map(i=>{ const n=selQ[i]; const share=kgVal!=null?(est>0?kgVal*(kgU(i)*n)/est:kgVal/Object.keys(selQ).length):kgU(i)*n; return {i,n,kg:Math.round(share*1000)/1000}; });
    const tok=moved.map(x=>unOf(x.i)>1?`${x.i}x${x.n}`:x.i).join(", ");
    await guard(async()=>{
      // 1) lote de destino
      let dl=destLot(); let destId;
      if(dl){ destId=dl.id; const d0=idsOf(dl); const ids=[...d0.ids], qq={...d0.q}, kg={...d0.kg};
        for(const x of moved){ if(ids.includes(x.i)){ qq[x.i]=((qq[x.i])||1)+x.n; kg[x.i]=Math.round(((Number(kg[x.i])||0)+x.kg)*1000)/1000; } else { ids.push(x.i); if(x.n>1) qq[x.i]=x.n; if(x.kg) kg[x.i]=x.kg; } }
        if(moved.length) await store.updateAmostra(destId,{ids,idsQ:qq,idsKg:kg,idsConf:d0.conf||"",atualizadoEm:data}); }
      else { const ids=moved.map(x=>x.i), qq={}, kg={}; moved.forEach(x=>{ if(x.n>1) qq[x.i]=x.n; if(x.kg) kg[x.i]=x.kg; });
        destId=await store.addAmostra({produto:dest,mapa:mapaDe(dest),codigo:"",lote:r.lote,tipo:r.tipo||"",conteudo:r.conteudo||"",fabricacao:r.fabricacao||"",validade:r.validade||"",unidade:r.unidade||"un",local:r.local||"",responsavel:resp,
          obs:`Reclassificado de ${r.produto} (lote ${r.lote}) em ${fmtD(data)}.`,origemReclass:r.id,...(moved.length?{ids,idsQ:qq,idsKg:kg}:{})},0,resp,data); }
      // 2) movimentações ligadas
      const base={lote:r.lote,qtd,data,motivo:mot,responsavel:resp,ids:tok,kg:kgMov,...(cli?{clienteId:cli.id,cliente:cli.nome,clienteReclass:true}:{})};
      await store.addMov({...base,amostraId:r.id,produto:r.produto,tipo:"reclass_out",destino:dest,reclassPara:destId,obs:`Para ${dest}${obs?". "+obs:""}`});
      await store.addMov({...base,amostraId:destId,produto:dest,tipo:"reclass_in",destino:"",reclassDe:r.id,obs:`De ${r.produto}${obs?". "+obs:""}`});
      // 3) tira os volumes da origem
      if(moved.length){ let ids=[...io.ids], qq={...io.q}, kg={...io.kg};
        for(const x of moved){ const un=unOf(x.i); if(x.n>=un){ ids=ids.filter(y=>y!==x.i); delete qq[x.i]; delete kg[x.i]; } else { const rest=un-x.n; if(rest>1) qq[x.i]=rest; else delete qq[x.i]; kg[x.i]=Math.max(0,Math.round(((Number(kg[x.i])||0)-x.kg)*1000)/1000); } }
        await store.updateAmostra(r.id,{ids,idsQ:qq,idsKg:kg,idsConf:io.conf||"",atualizadoEm:data}); }
      else await store.updateAmostra(r.id,{atualizadoEm:data});
    },"Reclassificação registrada");
  };
}

function ajustarLote(id){
  if(S.readonly||!FULL()) return;
  const r=rows().find(x=>x.id===id); if(!r) return;
  const io=idsOf(r);
  let L=io.ids.map(i=>({id:i,n:(io.q&&io.q[i])||1,kg:Number(io.kg[i])||0}));
  const fmt=v=>v?String(Math.round(v*1000)/1000).replace(".",","):"";
  const body=`<div class="ref full"><div><div class="prod">${esc(r.produto)}</div><div class="sub">Lote <span class="lote">${esc(r.lote)}</span> · saldo atual <b>${nf(r.saldo)}</b> ${esc(r.unidade||"")}</div></div></div>
   ${io.conf?`<div class="banner full" style="border-color:var(--warn);color:var(--warn);background:var(--warn-soft)">Da importação: ${esc(io.conf)}</div>`:""}
   <div class="full"><div class="sub" style="margin-bottom:6px">Corrija o ID, o número de volumes e o peso total de cada ID. O peso aceita vírgula e pode ser colado.</div>
     <div class="tablebox"><table><thead><tr><th>ID</th><th class="num">Volumes</th><th class="num">Peso total (kg)</th><th class="num">Peso por volume</th><th></th></tr></thead><tbody id="aj_rows"></tbody></table></div>
     <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap"><button type="button" class="btn sm" id="aj_add">+ Incluir ID</button></div></div>
   <div class="field"><label for="aj_saldo">Saldo do lote (unidades)</label><input id="aj_saldo" type="number" min="0" step="1" value="${r.saldo}"></div>
   <div class="field" style="justify-content:flex-end"><div class="sub" id="aj_sum"></div><button type="button" class="btn sm" id="aj_eq" style="align-self:flex-start;margin-top:4px">Igualar saldo à soma dos volumes</button></div>
   <div class="field full"><label for="aj_resp">Responsável *</label><select id="aj_resp">${opts(S.cfg.analistas,lastResp(),"Selecione")}</select></div>
   <div class="field full"><label for="aj_mot">Motivo do ajuste *</label><input id="aj_mot" value="Correção de dados importados da planilha"></div>
   <div class="formerr full" id="aj_err"></div>`;
  const m=modal("Ajustar volumes e saldo",body,`<button type="button" class="btn" data-close>Cancelar</button><button type="submit" class="btn primary">Salvar ajuste</button>`,true);
  const q=s=>m.ov.querySelector(s);
  const draw=()=>{
    q("#aj_rows").innerHTML=L.map((x,k)=>`<tr><td><input class="aj_id" data-k="${k}" value="${esc(x.id)}" style="font-family:var(--mono);width:100px"></td>
      <td class="num"><input class="aj_n" data-k="${k}" type="number" min="1" step="1" value="${x.n}" style="width:80px;text-align:right"></td>
      <td class="num"><input class="aj_kg" data-k="${k}" inputmode="decimal" value="${fmt(x.kg)}" placeholder="0,000" style="width:110px;text-align:right;font-family:var(--mono)"></td>
      <td class="num sub aj_pu" data-k="${k}">${x.kg&&x.n?fkg(x.kg/x.n)+" kg":"—"}</td>
      <td><button type="button" class="btn sm aj_rm" data-k="${k}" aria-label="Remover ${esc(x.id)}">Remover</button></td></tr>`).join("")||`<tr><td colspan="5" class="muted">Nenhum ID. Use “+ Incluir ID”.</td></tr>`;
    sum(); };
  const sum=()=>{ const t=L.reduce((s,x)=>s+(Number(x.n)||0),0), k=L.reduce((s,x)=>s+(Number(x.kg)||0),0), sd=Number(q("#aj_saldo").value)||0;
    q("#aj_sum").innerHTML=`Soma dos IDs: <b>${t}</b> volume${t===1?"":"s"} · ${fkg(k)} kg${t!==sd?` · <span style="color:var(--warn)">diferente do saldo (${sd})</span>`:` · <span style="color:var(--ok)">bate com o saldo</span>`}`; };
  m.ov.addEventListener("input",ev=>{ const k=+ev.target.dataset.k; const t=ev.target;
    if(t.classList.contains("aj_id")) L[k].id=t.value.trim().toUpperCase();
    else if(t.classList.contains("aj_n")) L[k].n=parseInt(t.value)||0;
    else if(t.classList.contains("aj_kg")) L[k].kg=parseKg(t.value)||0;
    if(t.classList.contains("aj_n")||t.classList.contains("aj_kg")){ const c=m.ov.querySelector(`.aj_pu[data-k="${k}"]`); if(c) c.textContent=L[k].kg&&L[k].n?fkg(L[k].kg/L[k].n)+" kg":"—"; }
    sum(); });
  m.ov.addEventListener("paste",ev=>{ if(!ev.target.classList.contains("aj_kg")) return; const t=(ev.clipboardData||window.clipboardData)?.getData("text"); if(!t) return; ev.preventDefault(); const v=parseKg(t); ev.target.value=v==null?t.trim():fmt(v); ev.target.dispatchEvent(new Event("input",{bubbles:true})); });
  m.ov.addEventListener("click",ev=>{ const b=ev.target.closest(".aj_rm"); if(b){ L.splice(+b.dataset.k,1); draw(); } });
  q("#aj_add").onclick=()=>{ L.push({id:"",n:1,kg:0}); draw(); const ins=m.ov.querySelectorAll(".aj_id"); ins[ins.length-1].focus(); };
  q("#aj_eq").onclick=()=>{ q("#aj_saldo").value=L.reduce((s,x)=>s+(Number(x.n)||0),0); sum(); };
  draw();
  m.form.onsubmit=async e=>{ e.preventDefault(); const err=q("#aj_err");
    const resp=q("#aj_resp").value, mot=q("#aj_mot").value.trim(), novo=Number(q("#aj_saldo").value);
    const names=L.map(x=>x.id);
    if(names.some(x=>!x)){ err.textContent="Há um ID sem nome. Preencha ou remova a linha."; return; }
    if(new Set(names).size!==names.length){ err.textContent="Há IDs repetidos. Junte os volumes numa linha só."; return; }
    if(L.some(x=>!(x.n>=1))){ err.textContent="Cada ID precisa ter pelo menos 1 volume."; return; }
    if(!(novo>=0)||Math.round(novo)!==novo){ err.textContent="O saldo precisa ser um número inteiro, zero ou maior."; return; }
    if(!resp){ err.textContent="Selecione o responsável pelo ajuste."; return; }
    if(!mot){ err.textContent="Informe o motivo do ajuste."; return; }
    LS.set("resp",resp); m.close();
    const ids=L.map(x=>x.id), qq={}, kg={}; L.forEach(x=>{ if(x.n>1) qq[x.id]=x.n; if(x.kg) kg[x.id]=Math.round(x.kg*1000)/1000; });
    const t=L.reduce((s,x)=>s+x.n,0);
    const conf=t===novo?"":`Após ajuste: os IDs somam ${t} volume(s); o saldo do lote é ${novo}.`;
    const before=io.ids.map(i=>idLabel(i,io.q,io.kg)).join(", ")||"nenhum", after=ids.map(i=>idLabel(i,qq,kg)).join(", ")||"nenhum";
    const ok=await guard(()=>store.updateAmostra(r.id,{ids,idsQ:qq,idsKg:kg,idsConf:conf,atualizadoEm:hoje()}));
    if(!ok) return;
    const diff=novo-r.saldo;
    const idsChanged=before!==after;
    if(diff!==0||idsChanged) await guard(()=>store.addMov({amostraId:r.id,produto:r.produto,lote:r.lote,tipo:"ajuste",qtd:diff,data:hoje(),motivo:mot,destino:"",responsavel:resp,
      obs:(diff!==0?`Saldo ${r.saldo} → ${novo}. `:"")+(idsChanged?`IDs antes: ${before}. Depois: ${after}.`:"").slice(0,1500),ids:"",kg:0}),"Ajuste salvo");
    else toast("Nada mudou.");
  };
}

function historico(id){
  const r=rows().find(x=>x.id===id); if(!r) return;
  const M=S.movs.filter(m=>m.amostraId===id).sort((a,b)=>String(b.data).localeCompare(String(a.data))||String(b.criadoEm).localeCompare(String(a.criadoEm)));
  const info=[["Registro MAPA",r.mapa],["Código interno",r.codigo],["Tipo",r.tipo],["Conteúdo",r.conteudo],["Fabricação",fmtD(r.fabricacao)],["Validade",fmtD(r.validade)],["Localização",r.local],["Cadastrado por",r.responsavel],["Observação",r.obs]].filter(x=>x[1]&&x[1]!=="—");
  const io=idsOf(r);
  const body=`<div class="ref full"><div><div class="prod">${esc(r.produto)}</div><div class="sub">Lote <span class="lote">${esc(r.lote)}</span></div></div><div>${pill(r.st)} <b style="font-family:var(--mono)">${nf(r.saldo)}</b> <span class="sub">${esc(r.unidade||"")}</span></div></div>
   <div class="full"><div class="sub" style="margin-bottom:4px">IDs em estoque (${io.ids.length})${io.ids.length?` · peso ${fkg(kgOf(r))} kg`:""}${io.conf?` · <span style="color:var(--warn)">conferir: ${esc(io.conf)}</span>`:""}</div>${io.ids.length?`<div class="chips">${io.ids.map(i=>`<span class="chip" style="padding:3px 10px;font-family:var(--mono)">${esc(idLabel(i,io.q,io.kg))}</span>`).join("")}</div>`:`<span class="sub">Nenhum ID registrado.</span>`}</div>
   <div class="full" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px">${info.map(([k,v])=>`<div><div class="sub">${k}</div><div>${esc(v)}</div></div>`).join("")}</div>
   <div class="full tablebox"><table><thead><tr><th>Data</th><th>Tipo</th><th class="num">Qtd</th><th>Motivo</th><th>Destino</th><th>Responsável</th><th>Obs.</th></tr></thead><tbody>
   ${M.map(m=>`<tr><td class="mono">${fmtD(m.data)}</td><td class="mv-${m.tipo}"><b>${MV_LABEL[m.tipo]||esc(m.tipo)}</b></td><td class="num">${Number(m.qtd)===0?"":((m.tipo==="entrada"||m.tipo==="reclass_in"||(m.tipo==="ajuste"&&m.qtd>0))?"+":"−")}${nf(Math.abs(m.qtd))}</td><td>${esc(m.motivo)}</td><td>${esc(m.destino)}</td><td>${esc(m.responsavel)}</td><td class="sub">${esc(m.obs)}${m.ids?` · ID ${esc(m.ids)}`:""}</td></tr>`).join("")||`<tr><td colspan="7" class="muted">Sem movimentações.</td></tr>`}
   </tbody></table></div>`;
  const ro=S.readonly;
  const m=modal("Detalhes da amostra",body,`${ro||!FULL()?"":`<button type="button" class="btn" id="h_ed" style="margin-right:auto">Editar cadastro</button>`}${!ro&&FULL()&&r.saldo>0&&!r.descartada?`<button type="button" class="btn" id="h_rc">Reclassificar</button>`:""}${!ro&&FULL()&&!r.descartada?`<button type="button" class="btn" id="h_aj">Ajustar volumes e saldo</button>`:""}${!ro&&r.saldo>0&&!r.descartada?`<button type="button" class="btn" id="h_rt">Retirar</button>`:""}<button type="button" class="btn primary" data-close>Fechar</button>`,true);
  const ed=m.ov.querySelector("#h_ed"); if(ed) ed.onclick=()=>{ m.close(); formAmostra(S.amostras.find(x=>x.id===id)); };
  const rt=m.ov.querySelector("#h_rt"); if(rt) rt.onclick=()=>{ m.close(); formMov(id,"saida"); };
  const aj=m.ov.querySelector("#h_aj"); if(aj) aj.onclick=()=>{ m.close(); ajustarLote(id); };
  const rc=m.ov.querySelector("#h_rc"); if(rc) rc.onclick=()=>{ m.close(); reclassificar(id); };
}

/* ======================= Importar ======================= */
let IMP=[];
const HMAP=[["produto",["produto","material","produto acabado","descricao","item","insumo","nome comercial"]],["mapa",["registro mapa","mapa","registro","n mapa","numero mapa","registro no mapa"]],["codigo",["codigo","cod","codigo do produto","codigo interno","sku"]],["lote",["lote","n lote","numero do lote"]],["qtd",["quantidade","qtd","qtde","quant"]],["fabricacao",["fabricacao","data de fabricacao","data fabricacao","fab"]],["validade",["validade","data de validade","vencimento","data validade","val"]],["local",["localizacao","local","localização"]],["tipo",["tipo","tipo de amostra"]],["unidade",["unidade","un"]],["responsavel",["responsavel","analista"]],["obs",["observacao","obs","observacoes"]]];
function toISO(v){
  if(v==null||v==="") return "";
  if(v instanceof Date) return isoLocal(new Date(v.getTime()+v.getTimezoneOffset()*0));
  if(typeof v==="number"){ const d=new Date(Math.round((v-25569)*864e5)); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`; }
  const s=String(v).trim(); let m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if(m){ let y=+m[3]; if(y<100) y+=2000; return `${y}-${pad(m[2])}-${pad(m[1])}`; }
  m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return `${m[1]}-${m[2]}-${m[3]}`;
  return "";
}
async function lerImport(file){
  const box=$("#impprev"); if(!file){ box.innerHTML=""; return; }
  if(!window.XLSX){ box.innerHTML=`<p class="formerr">O leitor de planilhas não carregou. Recarregue a página e tente de novo.</p>`; return; }
  try{
    const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{type:"array",cellDates:false});
    const ws=wb.Sheets[wb.SheetNames[0]]; const A=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});
    let hi=A.findIndex(r=>r.some(c=>["produto","material","lote"].includes(norm(c))));
    if(hi<0){ box.innerHTML=`<p class="formerr">Não encontrei um cabeçalho com “Produto” ou “Lote” na primeira aba.</p>`; return; }
    const H=A[hi].map(norm); const col={};
    for(const [k,names] of HMAP){ const i=H.findIndex(h=>names.includes(h)); if(i>=0) col[k]=i; }
    if(col.produto==null||col.lote==null){ box.innerHTML=`<p class="formerr">A planilha precisa ter as colunas Produto e Lote.</p>`; return; }
    IMP=A.slice(hi+1).map(r=>{ const g=k=>col[k]==null?"":r[col[k]]; return {produto:String(g("produto")).trim(),mapa:String(g("mapa")).trim()||mapaDe(String(g("produto")).trim()),codigo:String(g("codigo")).trim(),lote:String(g("lote")).trim(),qtd:Number(String(g("qtd")).replace(",","."))||0,fabricacao:toISO(g("fabricacao")),validade:toISO(g("validade")),local:String(g("local")).trim(),tipo:String(g("tipo")).trim()||S.cfg.tipos[0],unidade:String(g("unidade")).trim()||S.cfg.unidades[0],responsavel:String(g("responsavel")).trim(),obs:String(g("obs")).trim()}; })
      .filter(x=>x.produto&&x.lote);
    const semVal=IMP.filter(x=>!x.validade).length;
    box.innerHTML=`<p style="margin:0"><b>${IMP.length}</b> lote${IMP.length===1?"":"s"} encontrado${IMP.length===1?"":"s"}${semVal?` (${semVal} sem data de validade)`:""}. Colunas reconhecidas: ${Object.keys(col).join(", ")}.</p>
     <div class="tablebox"><table><thead><tr><th>Produto</th><th>Lote</th><th class="num">Qtd</th><th>Fabricação</th><th>Validade</th><th>Local</th></tr></thead><tbody>${IMP.slice(0,8).map(x=>`<tr><td>${esc(x.produto)}</td><td class="mono">${esc(x.lote)}</td><td class="num">${nf(x.qtd)}</td><td class="mono">${fmtD(x.fabricacao)}</td><td class="mono">${fmtD(x.validade)}</td><td>${esc(x.local)}</td></tr>`).join("")}</tbody></table></div>
     ${IMP.length>8?`<p class="sub" style="margin:0">… e mais ${IMP.length-8}.</p>`:""}
     <div><button class="btn primary" data-act="doimport"${IMP.length?"":" disabled"}>Importar ${IMP.length} lote${IMP.length===1?"":"s"}</button></div>`;
  }catch(e){ console.error(e); box.innerHTML=`<p class="formerr">Não consegui ler esse arquivo. Confira se é .xlsx ou .csv.</p>`; }
}
async function doImport(){
  const box=$("#impprev"); const list=IMP.slice(); IMP=[]; let n=0;
  box.innerHTML=`<p style="margin:0">Importando… 0 de ${list.length}</p>`;
  for(const x of list){
    const {qtd,...a}=x;
    const ok=await guard(()=>store.addAmostra(a,qtd,a.responsavel,hoje()));
    if(!ok) break; n++; const p=$("#impprev"); if(p) p.innerHTML=`<p style="margin:0">Importando… ${n} de ${list.length}</p>`;
  }
  const p=$("#impprev"); if(p) p.innerHTML=`<p style="margin:0"><b>${n}</b> lote${n===1?"":"s"} importado${n===1?"":"s"}.</p>`;
  const f=$("#impfile"); if(f) f.value="";
  toast(`${n} lotes importados`);
}

/* ======================= Exportar ======================= */
async function exportar(){
  if(!window.XLSX){ toast("O gerador de Excel não carregou. Recarregue a página."); return; }
  const R=rows().sort((a,b)=>a.produto.localeCompare(b.produto,"pt"));
  const est=R.map(r=>({"Produto":r.produto,"Registro MAPA":r.mapa||"","Código interno":r.codigo||"","Lote":r.lote,"Tipo":r.tipo||"","Conteúdo/unid.":r.conteudo||"","Fabricação":fmtD(r.fabricacao),"Validade":fmtD(r.validade),"Dias p/ vencer":r.dias??"","Situação":ST_LABEL[r.st],"Saldo":r.saldo,"Unidade":r.unidade||"","Localização":r.local||"","Peso em estoque (kg)":r.saldo>0&&!r.descartada?Math.round(kgOf(r)*1000)/1000:0,"IDs em estoque":idsOf(r).ids.map(i=>idLabel(i,idsOf(r).q,idsOf(r).kg)).join(", "),"Última atualização":fmtD(r.atual),"Responsável":r.respUlt,"Observação":r.obs||""}));
  const mv=S.movs.slice().sort((a,b)=>String(a.data).localeCompare(String(b.data))).map(m=>({"Data":fmtD(m.data),"Mês":m.data?`${MES[+m.data.slice(5,7)-1]}${m.data.slice(0,4)}`:"","Tipo":MV_LABEL[m.tipo]||m.tipo,"Produto":m.produto,"Lote":m.lote,"Quantidade":m.qtd,"Tipo de saída":(m.tipo==="saida"||m.tipo==="descarte")?catOf(m):"","Motivo":m.motivo||"","Destino":m.cliente||m.destino||"","ID(s)":m.ids||"","Peso (kg)":m.kg||"","Responsável":m.responsavel||"","Observação":m.obs||"","Fotos":m.fotos||0}));
  const wb=XLSX.utils.book_new();
  const add=(name,data,cols)=>{ const ws=XLSX.utils.json_to_sheet(data.length?data:[{"":"Sem registros"}]); ws["!cols"]=cols.map(w=>({wch:w})); XLSX.utils.book_append_sheet(wb,ws,name); };
  const cw=[44,20,12,14,20,14,12,12,10,12,8,8,22,14,16,30];
  add("Estoque",est,cw);
  add("Vencidas",est.filter(x=>x["Situação"]==="Vencida"),cw);
  const cn=id=>{ const c=S.clientes.find(x=>x.id===id); return c?c:{}; };
  if(S.clientes.length) add("Clientes e laboratórios",S.clientes.slice().sort((a,b)=>a.nome.localeCompare(b.nome,"pt")).map(c=>({"Nome":c.nome,"Tipo":c.tipo||"Cliente","País":c.pais||"","Cidade":c.cidade||"","Contato":c.contato||"","E-mail / telefone":c.email||"","Observação":c.obs||""})),[28,18,14,18,20,26,30]);
  const ev=enviosAll(); if(ev.length) add("Envios clientes-lab",ev.slice().sort((a,b)=>String(a.data).localeCompare(String(b.data))).map(e=>({"Data":fmtD(e.data),"Cliente / laboratório":cn(e.clienteId).nome||e.cliente||"","Tipo":cn(e.clienteId).tipo||"","País":cn(e.clienteId).pais||"","Produto":e.produto,"Lote":e.lote,"ID(s)":e.ids||"","Quantidade":e.qtd,"Peso (kg)":e.kg||"","Responsável":e.responsavel||"","Observação":e.obs||"","Origem":e.origem==="planilha"?"Planilha":"App"})),[12,26,14,50,12,10,16,30,10]);
  if(S.hist.length) add("Histórico mensal",S.hist.slice().sort((a,b)=>String(a.mes).localeCompare(b.mes)||String(a.produto).localeCompare(b.produto,"pt")).map(h=>({"Mês":h.mes,"Produto":h.produto,"Unidades retiradas":h.qtd,"Lotes":h.lotes||""})),[10,56,18,8]);
  add("Produtos",PRODS().map(p=>({"Produto":p.nome,"Registro MAPA":p.mapa||"","Categoria":p.categoria||""})),[56,22,14]);
  add("Movimentações",mv,[12,10,10,30,14,10,18,24,12,16,10,16,30,8]);
  const out=XLSX.write(wb,{bookType:"xlsx",type:"array"});
  const fname=`Amostras_Produto_Acabado_${hoje().split("-").reverse().join(".")}.xlsx`;
  const dlc=DL;
  try{ await dlc.save({filename:fname,data:new Uint8Array(out)}); toast("Planilha exportada"); }
  catch(e){ if(e&&e.code!=="declined") toast("Não foi possível exportar agora."); }
}

/* ---------- Backup completo (migração) ---------- */
async function backupJSON(btn){
  if(!S.isOwner||!db) return; const t0=btn.textContent; btn.disabled=true; btn.textContent="Preparando…";
  try{
    const strip=a=>a.map(x=>({...x}));
    const out={app:"amostras-pa",versao:1,geradoEm:new Date().toISOString(),
      colecoes:{amostras:strip(S.amostras),movs:strip(S.movs),clientes:strip(S.clientes),historico:strip(S.hist),envios_hist:strip(S.envh)},
      config:{listas:{...S.cfg},estoque_ids:{map:S.stkIds},estoque_kg:{map:S.stkKg},envios_ids:{map:S.envIds},envios_kg:{map:S.envKg}},
      pessoas:[], log:[]};
    let names={}; try{ names=await me.profiles(S.pessoas.map(p=>p.id)); }catch(e){}
    out.pessoas=S.pessoas.filter(p=>!p.origem).map(p=>({id:p.id,nome:p.nome||p.email||"",primeiroAcesso:p.primeiroAcesso||"",dono:!!p.dono,nivel:p.dono?"proprietária":(S.acessos[p.id]||"pendente")}));
    for(const p of S.pessoas){ try{ const q=await db.doc("log/"+p.id).collection("meses").get(); q.docs.forEach(d=>out.log.push({uid:p.id,mes:d.id,eventos:d.data().eventos||[]})); }catch(e){} }
    const txt=JSON.stringify(out);
    const dlc=DL;
    await dlc.save({filename:`backup_amostras_${hoje()}.json`,data:txt});
    toast(`Backup gerado: ${out.colecoes.amostras.length} amostras, ${out.colecoes.movs.length} movimentações`);
  }catch(e){ if(!(e&&e.code==="declined")){ console.error(e); toast("Não foi possível gerar o backup."); } }
  finally{ btn.disabled=false; btn.textContent=t0; }
}

/* ======================= Tela de entrada ======================= */
const WEL={ greet(hh){ const h=hh??new Date().getHours(); return h<12?"Bom dia":h<18?"Boa tarde":"Boa noite"; },
  sky(hh){ const h=hh??new Date().getHours(); const W=document.getElementById("welcome"); if(!W) return;
    const set=(a,b,sun,glow,sunc)=>{ document.getElementById("wsky1").setAttribute("stop-color",a); document.getElementById("wsky2").setAttribute("stop-color",b); const s=document.getElementById("wsun"); s.setAttribute("cy",sun); s.setAttribute("fill",sunc); const g=document.getElementById("wsunglow"); g.setAttribute("cy",sun); g.style.opacity=glow; document.getElementById("wnight").setAttribute("opacity", glow<.5?".38":"0"); };
    if(h<6||h>=19) set("#1d2a4a","#4a5d7a",130,.25,"#f2f0e6"); else if(h<9) set("#f6c9a8","#f3e6d6",240,1,"#ffc76b"); else if(h>=17) set("#f2b38c","#f6e2cf",250,1,"#ffb35c"); else set("#a9d4f0","#e6f2ec",150,1,"#ffd97a"); },
  async fill(){ const t=document.getElementById("wTitle"); let nome="";
    try{ if(me){ const v=await me.me(); nome=(v&&v.name||"").split(" ")[0]; } }catch(e){}
    t.textContent=`${WEL.greet()}${nome?", "+nome:""}!`;
    const r=document.getElementById("wRole"); if(S.mode==="db"&&AUTH.ok!==undefined){ r.textContent=ROLE_LABEL[S.role]+(S.isOwner?" · proprietária":""); r.hidden=false; } },
  close(){ const W=document.getElementById("welcome"); if(!W||W.classList.contains("out")||!AUTH.ok) return; W.classList.add("out"); setTimeout(()=>{ W.hidden=true; },500); const f=document.querySelector("button.nv[aria-selected=true]"); f&&f.focus({preventScroll:true}); } };
WEL.sky(); document.getElementById("wTitle").textContent=WEL.greet()+"!";
document.getElementById("wPeek").addEventListener("click",e=>{ const W=document.getElementById("welcome"); const on=W.classList.toggle("peek"); e.currentTarget.setAttribute("aria-pressed",String(on)); e.currentTarget.querySelector("span").textContent=on?"Mostrar entrada":"Ver a paisagem"; });

/* ---------- Login (Firebase), sessão e histórico de acesso ---------- */
const AUTH={ ok:false };
const UA=()=>{ const u=navigator.userAgent; const os=/Android/i.test(u)?"Android":/iPhone|iPad/i.test(u)?"iPhone/iPad":/Windows/i.test(u)?"Windows":/Mac OS/i.test(u)?"Mac":/Linux/i.test(u)?"Linux":"Outro"; const br=/Edg\//.test(u)?"Edge":/OPR\//.test(u)?"Opera":/Chrome\//.test(u)?"Chrome":/Firefox\//.test(u)?"Firefox":/Safari\//.test(u)?"Safari":""; const pwa=window.matchMedia&&window.matchMedia("(display-mode: standalone)").matches?" · app instalado":""; return os+(br?" · "+br:"")+pwa; };
const EV_LABEL={entrada:"Entrada",saida:"Saída",cadastro:"Conta criada",senha_redefinida:"Senha redefinida",reset_enviado:"E-mail de nova senha enviado",sessao:"Sessão expirada (8 h)",falha:"Senha incorreta",senha_criada:"Senha criada",bloqueio:"Bloqueio por tentativas"};
async function logEv(ev,uid,extra){ if(!db) return; const id=uid||S.myId; if(!id) return;
  const ref=db.doc("log/"+id).collection("meses").doc(hoje().slice(0,7));
  try{ const d=await ref.get(); const evs=d.exists?(d.data().eventos||[]):[]; evs.push({t:new Date().toISOString(),ev,ua:UA(),...(extra||{})}); await ref.set({eventos:evs.slice(-3000)}); }catch(e){ console.error(e); } }
const SESS_H=8;
const loginKey=u=>"apa:login:"+u;
let PENDING_LOGIN=false, SUBS=[], DATA_ON=false;
function unsubAll(){ SUBS.forEach(f=>{ try{ f(); }catch(e){} }); SUBS=[]; DATA_ON=false; }
function wBody(html){ document.getElementById("wBody").innerHTML=html; const f=document.querySelector("#wBody input,#wBody button"); f&&setTimeout(()=>f.focus(),40); }
function wErr(t,ok){ const e=document.getElementById("wErr"); if(e){ e.textContent=t; e.style.color=ok?"#2f7a4f":""; } }
function showWelcome(){ const W=document.getElementById("welcome"); W.hidden=false; W.classList.remove("out","peek"); WEL.sky(); }
let W_MODE="";
const lastEmail=()=>{ try{ return localStorage.getItem("apa:email")||""; }catch(e){ return ""; } };
const AUTH_ERR=c=>({"auth/invalid-credential":"E-mail ou senha incorretos.","auth/wrong-password":"E-mail ou senha incorretos.","auth/user-not-found":"E-mail ou senha incorretos.","auth/invalid-login-credentials":"E-mail ou senha incorretos.",
  "auth/too-many-requests":"Muitas tentativas seguidas. Aguarde alguns minutos ou use “Esqueci minha senha”.","auth/email-already-in-use":"Já existe uma conta com esse e-mail. Use “Entrar”.","auth/invalid-email":"E-mail inválido.",
  "auth/weak-password":"A senha precisa ter pelo menos 6 caracteres.","auth/network-request-failed":"Sem conexão com a internet. Tente de novo.","auth/user-disabled":"Esta conta foi desativada."})[c]||"Não foi possível concluir agora. Tente de novo.";
function showLogin(msg){ W_MODE="login"; wBody(`${msg?`<p class="wmsg" style="margin-bottom:12px">${msg}</p>`:""}<div class="wf"><label for="wEm">E-mail</label><input id="wEm" type="email" autocomplete="username" value="${esc(lastEmail())}" required></div>
  <div class="wf"><label for="wPw">Senha</label><input id="wPw" type="password" autocomplete="current-password" required></div><div class="werr" id="wErr" role="alert"></div>
  <button type="submit" class="wgo">Entrar</button><div class="wlinks"><button type="button" class="wlink" data-w="signup">Criar conta</button><button type="button" class="wlink" data-w="reset">Esqueci minha senha</button></div>`);
  if(lastEmail()) setTimeout(()=>{ const p=document.getElementById("wPw"); p&&p.focus(); },60); }
function showSignup(){ W_MODE="signup"; wBody(`<p class="wmsg" style="margin-bottom:12px">Crie a sua conta. Depois de confirmar o e-mail, a proprietária libera o seu nível de acesso.</p>
  <div class="wf"><label for="wNm">Nome</label><input id="wNm" autocomplete="name" required></div>
  <div class="wf"><label for="wEm">E-mail</label><input id="wEm" type="email" autocomplete="username" required></div>
  <div class="wf"><label for="wPw">Senha (mínimo 6 caracteres)</label><input id="wPw" type="password" autocomplete="new-password" minlength="6" required></div>
  <div class="wf"><label for="wPw2">Repita a senha</label><input id="wPw2" type="password" autocomplete="new-password" required></div><div class="werr" id="wErr" role="alert"></div>
  <button type="submit" class="wgo">Criar conta</button><div class="wlinks"><button type="button" class="wlink" data-w="login">Já tenho conta</button></div>`); }
function showReset(){ W_MODE="reset"; wBody(`<p class="wmsg" style="margin-bottom:12px">Informe o seu e-mail. Você recebe um link para criar uma senha nova.</p>
  <div class="wf"><label for="wEm">E-mail</label><input id="wEm" type="email" autocomplete="username" value="${esc(lastEmail())}" required></div><div class="werr" id="wErr" role="alert"></div>
  <button type="submit" class="wgo">Enviar link</button><div class="wlinks"><button type="button" class="wlink" data-w="login">Voltar</button></div>`); }
function showVerify(email){ W_MODE="verify"; wBody(`<p class="wmsg" style="margin-bottom:12px">Enviamos um link de confirmação para <b>${esc(email)}</b>. Abra o e-mail (veja também o spam), clique no link e volte aqui.</p><div class="werr" id="wErr" role="alert"></div>
  <button type="submit" class="wgo">Já confirmei</button><div class="wlinks"><button type="button" class="wlink" data-w="resend">Reenviar e-mail</button><button type="button" class="wlink" data-w="out">Sair</button></div>`); }
function showMsg(html){ W_MODE="msg"; AUTH.ok=false; showWelcome(); wBody(`<p class="wmsg">${html}</p><div class="wlinks"><button type="button" class="wlink" data-w="out">Sair</button></div>`); }
const showBlocked=()=>showMsg("O seu acesso a este app está bloqueado. Fale com a proprietária do controle.");
const showPending=()=>showMsg("Conta confirmada. Agora a proprietária precisa liberar o seu nível de acesso em <b>Cadastros → Acessos</b>. Depois disso, é só abrir o app de novo.");

WEL.start=function(){ /* chamado pelo fluxo antigo; o Firebase decide em onUser/gate */ };
function gate(){
  if(S.mode!=="db") return;
  if(S.role==="bloqueado"){ unsubData(); return showBlocked(); }
  if(S.role==="pendente"){ unsubData(); return showPending(); }
  subscribeData(); AUTH.ok=true; WEL.close();
}
let DATA_SUBS=[];
function unsubData(){ DATA_SUBS.forEach(f=>{ try{ f(); }catch(e){} }); DATA_SUBS=[]; DATA_ON=false; }
function subscribeData(){
  if(DATA_ON) return; DATA_ON=true;
  const onErr=e=>{ console.error(e); if(e&&e.code==="permission-denied") return; S.mode="error"; render(); };
  const col=(c,fn)=>DATA_SUBS.push(db.collection(c).onSnapshot(s=>{ fn(s.docs.map(d=>({id:d.id,...d.data()}))); refresh(); },onErr));
  const doc=(p,fn)=>DATA_SUBS.push(db.doc(p).onSnapshot(s=>{ fn(s.exists?s.data():null); refresh(); },onErr));
  col("amostras",L=>S.amostras=L); col("movs",L=>S.movs=L); col("clientes",L=>S.clientes=L);
  col("pessoas",L=>S.pessoas=L);
  col("legado",L=>{ const pick=pre=>L.filter(d=>d.id.startsWith(pre)).sort((a,b)=>a.id.localeCompare(b.id,undefined,{numeric:true})).flatMap(d=>d.lista||[]);
    S.envh=pick("envios_"); S.hist=pick("historico_"); });
  doc("config/estoque_kg",d=>S.stkKg=(d&&d.map)||{}); doc("config/envios_kg",d=>S.envKg=(d&&d.map)||{});
  doc("config/estoque_ids",d=>S.stkIds=(d&&d.map)||{}); doc("config/envios_ids",d=>S.envIds=(d&&d.map)||{});
  doc("config/listas",d=>S.cfg=d?{...DEFAULT_CFG,...d}:{...DEFAULT_CFG});
}
async function onUser(u){
  unsubAll(); unsubData(); AUTH.ok=false;
  S.amostras=[]; S.movs=[]; S.clientes=[]; S.pessoas=[]; S.envh=[]; S.hist=[]; S.acessos={};
  document.querySelectorAll(".overlay").forEach(o=>o.remove());
  if(!u){ S.mode="auth"; S.myId=null; S.isOwner=false; me=null; render(); showWelcome(); WEL.fill(); document.getElementById("wRole").hidden=true; showLogin(); return; }
  const em=(u.email||"").toLowerCase();
  try{ localStorage.setItem("apa:email",em); }catch(e){}
  let t=0; try{ t=Number(localStorage.getItem(loginKey(u.uid))||0); }catch(e){}
  if(PENDING_LOGIN){ PENDING_LOGIN=false; t=Date.now(); try{ localStorage.setItem(loginKey(u.uid),String(t)); }catch(e){} u.__novo=true; }
  else if(!t || Date.now()-t>SESS_H*3600e3){ await firebase.auth().signOut(); return; }
  if(!u.emailVerified){ S.mode="auth"; showWelcome(); return showVerify(em); }
  S.myId=em; S.myName=u.displayName||""; S.isOwner = em===String(window.OWNER_EMAIL||"").toLowerCase();
  me={ me:async()=>({name:S.myName||S.myId}), profiles:async ids=>Object.fromEntries([].concat(ids).map(i=>{ const p=S.pessoas.find(x=>x.id===i); return [i,{name:p?(p.nome||p.email||i):i}]; })) };
  S.mode="db"; computeRole(); WEL.fill();
  try{ const ref=db.doc("pessoas/"+em); const d=await ref.get(); await ref.set({email:em,nome:S.myName||(d.exists&&d.data().nome)||"",ultimoAcesso:new Date().toISOString(),...(d.exists?{}:{primeiroAcesso:new Date().toISOString()}),...(S.isOwner?{dono:true}:{})},{merge:true}); }catch(e){ console.error(e); }
  if(u.__novo) logEv("entrada");
  const onAcc=()=>{ computeRole(); render(); WEL.fill(); gate(); };
  if(S.isOwner) SUBS.push(db.collection("niveis").onSnapshot(s=>{ S.acessos=Object.fromEntries(s.docs.map(d=>[d.id,d.data().nivel])); onAcc(); },e=>console.error(e)));
  else SUBS.push(db.doc("niveis/"+em).onSnapshot(s=>{ S.acessos=s.exists?{[em]:s.data().nivel}:{}; onAcc(); },e=>{ console.error(e); onAcc(); }));
}
document.getElementById("wForm").addEventListener("submit",async e=>{ e.preventDefault(); const btn=e.target.querySelector(".wgo"); if(btn&&btn.disabled) return;
  const v=id=>{ const el=document.getElementById(id); return el?el.value.trim():""; };
  const busy=(on,txt)=>{ if(!btn) return; btn.disabled=on; if(txt) btn.textContent=txt; };
  const auth=window.firebase&&firebase.auth&&firebase.apps.length?firebase.auth():null;
  if(W_MODE==="demo"){ AUTH.ok=true; WEL.close(); return; }
  if(!auth) return;
  try{
    if(W_MODE==="login"){ const em=v("wEm"), pw=document.getElementById("wPw").value; if(!em||!pw){ wErr("Preencha e-mail e senha."); return; }
      busy(true,"Entrando…"); PENDING_LOGIN=true; await auth.signInWithEmailAndPassword(em,pw); return; }
    if(W_MODE==="signup"){ const nm=v("wNm"), em=v("wEm"), a=document.getElementById("wPw").value, b2=document.getElementById("wPw2").value;
      if(!nm){ wErr("Informe o seu nome."); return; } if(a.length<6){ wErr("A senha precisa ter pelo menos 6 caracteres."); return; } if(a!==b2){ wErr("As duas senhas não são iguais."); return; }
      busy(true,"Criando…"); PENDING_LOGIN=true; const cr=await auth.createUserWithEmailAndPassword(em,a);
      await cr.user.updateProfile({displayName:nm}); try{ await cr.user.sendEmailVerification(); }catch(_){}
      try{ await db.doc("pessoas/"+em.toLowerCase()).set({email:em.toLowerCase(),nome:nm,primeiroAcesso:new Date().toISOString()},{merge:true}); }catch(_){}
      showVerify(em.toLowerCase()); return; }
    if(W_MODE==="reset"){ const em=v("wEm"); if(!em){ wErr("Informe o seu e-mail."); return; } busy(true,"Enviando…");
      await auth.sendPasswordResetEmail(em); busy(false,"Enviar link"); wErr("Pronto. Se houver uma conta com esse e-mail, o link chega em alguns minutos.",true); return; }
    if(W_MODE==="verify"){ const u=auth.currentUser; if(!u) return showLogin(); busy(true,"Verificando…"); await u.reload(); 
      if(!auth.currentUser.emailVerified){ busy(false,"Já confirmei"); wErr("Ainda não aparece como confirmado. Clique no link do e-mail e tente de novo."); return; }
      await auth.currentUser.getIdToken(true); PENDING_LOGIN=true; await logEv("cadastro",auth.currentUser.email.toLowerCase()).catch(()=>{}); onUser(auth.currentUser); return; }
  }catch(err){ console.error(err); PENDING_LOGIN=false; wErr(AUTH_ERR(err&&err.code)); busy(false, W_MODE==="login"?"Entrar":W_MODE==="signup"?"Criar conta":W_MODE==="reset"?"Enviar link":"Já confirmei"); }
});
document.getElementById("wForm").addEventListener("click",async e=>{ const w=e.target.closest("[data-w]"); if(!w) return; const k=w.dataset.w;
  if(k==="login") showLogin(); else if(k==="signup") showSignup(); else if(k==="reset") showReset();
  else if(k==="out"){ try{ await firebase.auth().signOut(); }catch(_){ } }
  else if(k==="resend"){ try{ await firebase.auth().currentUser.sendEmailVerification(); wErr("E-mail reenviado.",true); }catch(err){ wErr(AUTH_ERR(err&&err.code)); } }
});
async function sairApp(){ if(S.mode!=="db") return; await logEv("saida"); try{ const u=firebase.auth().currentUser; if(u) localStorage.removeItem(loginKey(u.uid)); }catch(e){} await firebase.auth().signOut(); }

/* ---------- Download de arquivos (Excel e backup) ---------- */
const DL={ async save({filename,data}){ const blob=data instanceof Blob?data:new Blob([data]); const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=filename; document.body.appendChild(a); a.click(); setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); },1500); return {status:"saved"}; } };

/* ---------- Instalar como app ---------- */
let INSTALL_EVT=null;
window.addEventListener("beforeinstallprompt",e=>{ e.preventDefault(); INSTALL_EVT=e; refresh(true); });
window.addEventListener("appinstalled",()=>{ INSTALL_EVT=null; toast("App instalado"); refresh(true); });
async function instalarApp(){ if(!INSTALL_EVT) return; INSTALL_EVT.prompt(); try{ await INSTALL_EVT.userChoice; }catch(e){} INSTALL_EVT=null; refresh(true); }
if("serviceWorker" in navigator && location.protocol.startsWith("http")) window.addEventListener("load",()=>navigator.serviceWorker.register("sw.js").catch(e=>console.warn(e)));

/* ---------- Importar backup do app do Claude ---------- */
let BKMSG="";
async function importarBackup(file){
  if(!file||!S.isOwner) return; let B;
  try{ B=JSON.parse(await file.text()); }catch(e){ toast("Arquivo inválido. Use o .json baixado no app do Claude."); return; }
  if(!B||B.app!=="amostras-pa"||!B.colecoes){ toast("Este arquivo não é um backup do app de amostras."); return; }
  const C=B.colecoes, n=(C.amostras||[]).length, m=(C.movs||[]).length;
  confirmar("Importar backup",`Serão gravados ${n} lotes, ${m} movimentações, ${(C.clientes||[]).length} clientes/laboratórios, o histórico da planilha, os IDs e pesos, as listas, ${(B.pessoas||[]).length} pessoas e o histórico de acesso. Registros com o mesmo código são substituídos.`,"Importar",async()=>{
    const say=t=>{ BKMSG=t; const box=$("#bkprev"); if(box) box.innerHTML=`<p style="margin:0">${t}</p>`; };
    const ops=[]; const put=(p,d)=>ops.push([p,d]);
    const clean=o=>{ const x={...o}; delete x.id; return x; };
    for(const c of ["amostras","movs","clientes"]) (C[c]||[]).forEach(d=>put(c+"/"+d.id,clean(d)));
    const chunk=(L,pre)=>{ for(let i=0;i*300<L.length;i++) put("legado/"+pre+i,{lista:L.slice(i*300,i*300+300)}); };
    chunk(C.envios_hist||[],"envios_"); chunk(C.historico||[],"historico_");
    const cf=B.config||{}; for(const k of ["listas","estoque_ids","estoque_kg","envios_ids","envios_kg"]) if(cf[k]) put("config/"+k,cf[k]);
    (B.pessoas||[]).forEach(p=>put("pessoas/claude:"+p.id,{nome:(p.nome||"Pessoa")+" (app Claude)",origem:"claude",primeiroAcesso:p.primeiroAcesso||"",nivelAntigo:p.nivel||""}));
    (B.log||[]).forEach(l=>put("log/claude:"+l.uid+"/meses/"+l.mes,{eventos:l.eventos||[]}));
    let done=0;
    try{ for(let i=0;i<ops.length;i+=400){ const bt=db.batch(); ops.slice(i,i+400).forEach(([p,d])=>bt.set(db.doc(p),d)); await bt.commit(); done=Math.min(ops.length,i+400); say(`Importando… ${done} de ${ops.length} registros`); }
      say(`<b>Backup importado.</b> ${ops.length} registros gravados. Agora libere o acesso da equipe em Acessos.`); toast("Backup importado");
    }catch(e){ console.error(e); say(`Parou em ${done} de ${ops.length}. Confira as regras do Firestore e tente de novo (o que já foi gravado é substituído sem duplicar).`); toast("Não foi possível concluir a importação."); }
    const f=$("#bkfile"); if(f) f.value="";
  });
}

["btnSair","btnSair2"].forEach(i=>{ const b=document.getElementById(i); if(b) b.addEventListener("click",sairApp); });
init();
</script>
</body>
</html>

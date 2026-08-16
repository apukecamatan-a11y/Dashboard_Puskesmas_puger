<!DOCTYPE html>
<html lang="id">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Papan Informasi — Puskesmas Puger</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link
    href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,400;0,500;0,600;0,700;0,800;1,500&family=IBM+Plex+Mono:wght@400;500;600&display=swap"
    rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
  <script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      --primary: #EC9CB6;
      --primary-dark: #C26A88;
      --primary-light: #FDEEF4;
      --accent: #F4B8CE;
      --accent-dark: #D98AA8;
      --bg: #FFF6F9;
      --surface: #FFFFFF;
      --text: #3A2530;
      --muted: #8C6B78;
      --border: #F3DDE6;
      --danger: #D64545;
      --radius: 16px;
      --shadow: 0 4px 24px rgba(214, 110, 150, 0.12);
      --font-display: 'Plus Jakarta Sans', sans-serif;
      --font-mono: 'IBM Plex Mono', monospace;
      --gap: 22px;
      --pad: 22px;
      --ticker-duration: 32s;
      /* warna grafik capaian berdasarkan level pencapaian (dapat diubah lewat Pengaturan Tampilan) */
      --chart-excellent: #4CAF7D;
      --chart-good: #E8A33D;
      --chart-poor: #D64545;
    }

    body.layout-compact {
      --gap: 14px;
      --pad: 14px;
    }

    * {
      box-sizing: border-box;
    }

    html,
    body {
      margin: 0;
      padding: 0;
    }

    body {
      font-family: var(--font-display);
      background: var(--bg);
      color: var(--text);
      -webkit-font-smoothing: antialiased;
    }

    a {
      color: inherit;
    }

    button {
      font-family: inherit;
      cursor: pointer;
    }

    input,
    select,
    textarea {
      font-family: inherit;
    }

    .hidden {
      display: none !important;
    }

    .container {
      max-width: 1180px;
      margin: 0 auto;
      padding: 0 24px;
    }

    /* ---------- TICKER ---------- */
    .ticker-bar {
      background: var(--primary-dark);
      color: #fff;
      overflow: hidden;
      white-space: nowrap;
      position: relative;
      font-size: 13.5px;
      letter-spacing: .2px;
    }

    .ticker-bar .ticker-track {
      display: inline-flex;
      padding-left: 100%;
      animation: ticker-scroll var(--ticker-duration, 32s) linear infinite;
    }

    .ticker-bar .ticker-track span {
      padding: 9px 40px;
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }

    .ticker-bar .ticker-track span::before {
      content: "●";
      color: var(--accent);
      font-size: 8px;
    }

    @keyframes ticker-scroll {
      0% {
        transform: translateX(0);
      }

      100% {
        transform: translateX(-100%);
      }
    }

    .ticker-empty {
      padding: 9px 24px;
      font-size: 13.5px;
      opacity: .75;
    }

    /* ---------- HERO ---------- */
    .hero {
      background: linear-gradient(160deg, var(--primary) 0%, var(--primary-dark) 100%);
      color: #fff;
      padding: 38px 0 64px;
      position: relative;
    }

    .hero-top {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 20px;
      flex-wrap: wrap;
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 16px;
      flex-wrap: wrap;
    }

    /* BESAR: foto profil Puskesmas/Dinas diperbesar & jadi lingkaran berlatar putih agar menonjol */
    .brand-mark {
      width: 78px;
      height: 78px;
      border-radius: 50%;
      background: #fff;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 30px;
      border: 3px solid rgba(255, 255, 255, .55);
      box-shadow: 0 6px 18px rgba(0, 0, 0, .18);
      flex-shrink: 0;
      padding: 5px;
      overflow: hidden;
    }

    .brand-mark img {
      width: 100%;
      height: 100%;
      object-fit: contain;
      border-radius: 50%;
    }

    .brand-text h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 800;
      letter-spacing: -.2px;
    }

    .brand-text p {
      margin: 3px 0 0;
      font-size: 12.5px;
      opacity: .9;
      font-weight: 500;
    }

    .hero-actions {
      display: flex;
      gap: 12px;
      align-items: center;
    }

    /* Logo Pemerintah Kabupaten (pemda) tampil di sisi kiri, berdampingan dengan logo profil */
    .pemda-logo-inline {
      height: 58px;
      width: auto;
      filter: drop-shadow(0 4px 10px rgba(0, 0, 0, .22));
      flex-shrink: 0;
    }

    .btn {
      border: none;
      border-radius: 10px;
      padding: 10px 16px;
      font-weight: 600;
      font-size: 13.5px;
      display: inline-flex;
      align-items: center;
      gap: 7px;
      transition: transform .12s ease, opacity .12s ease;
    }

    .btn:active {
      transform: scale(.97);
    }

    .btn-ghost {
      background: rgba(255, 255, 255, .14);
      color: #fff;
      border: 1px solid rgba(255, 255, 255, .3);
    }

    .btn-ghost:hover {
      background: rgba(255, 255, 255, .22);
    }

    .btn-accent {
      background: var(--accent);
      color: #3A2604;
    }

    .btn-accent:hover {
      opacity: .92;
    }

    .btn-primary {
      background: var(--primary);
      color: #fff;
    }

    .btn-primary:hover {
      background: var(--primary-dark);
    }

    .btn-danger {
      background: #FCEAEA;
      color: var(--danger);
    }

    .btn-danger:hover {
      background: #f8d6d6;
    }

    .btn-outline {
      background: transparent;
      border: 1px solid var(--border);
      color: var(--text);
    }

    .btn-outline:hover {
      background: #f2f5f3;
    }

    .btn-sm {
      padding: 7px 11px;
      font-size: 12.5px;
      border-radius: 8px;
    }

    .btn:disabled {
      opacity: .5;
      cursor: not-allowed;
    }

    /* Teks kontur hitam bergaris putih — dipakai untuk judul agar tidak polos/kaku */
    .txt-outline {
      color: #171314;
      text-shadow:
        -2px -2px 0 #fff, 2px -2px 0 #fff, -2px 2px 0 #fff, 2px 2px 0 #fff,
        0 -2.2px 0 #fff, 0 2.2px 0 #fff, -2.2px 0 0 #fff, 2.2px 0 0 #fff;
    }

    .txt-outline-lg {
      color: #171314;
      text-shadow:
        -3px -3px 0 #fff, 3px -3px 0 #fff, -3px 3px 0 #fff, 3px 3px 0 #fff,
        0 -3px 0 #fff, 0 3px 0 #fff, -3px 0 0 #fff, 3px 0 0 #fff,
        0 8px 16px rgba(58, 37, 48, .18);
    }

    .hero-headline {
      margin-top: 34px;
      max-width: 660px;
    }

    .hero-headline .eyebrow {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      font-family: var(--font-mono);
      font-size: 11.5px;
      letter-spacing: 1.3px;
      text-transform: uppercase;
      color: #fff;
      background: rgba(255, 255, 255, .2);
      border: 1px solid rgba(255, 255, 255, .4);
      padding: 7px 15px;
      border-radius: 999px;
      margin-bottom: 16px;
      transform: rotate(-1.2deg);
    }

    .hero-headline h2 {
      font-size: 34px;
      line-height: 1.28;
      margin: 0 0 12px;
      font-weight: 800;
      letter-spacing: -.4px;
    }

    /* kata kunci disorot seperti stabilo agar judul terasa lebih hidup, tidak seragam */
    .hero-headline h2 .hl-word {
      display: inline-block;
      color: var(--primary-dark);
      text-shadow: none;
      background: #fff;
      padding: 1px 12px;
      border-radius: 9px;
      transform: rotate(-1.5deg);
      margin: 0 2px;
      box-shadow: 0 5px 14px rgba(58, 37, 48, .22);
    }

    .hero-headline p {
      font-size: 15px;
      opacity: .92;
      margin: 0;
      line-height: 1.6;
      font-weight: 500;
    }

    .wave-divider {
      display: block;
      width: 100%;
      height: 54px;
      margin-top: -2px;
    }

    .wave-divider path {
      fill: var(--bg);
    }


    /* ---------- INFO NAV (pengelompokan informasi publik) ---------- */
    .info-nav {
      display: flex;
      gap: 8px;
      overflow-x: auto;
      padding: 16px 0 6px;
      margin-bottom: 8px;
      scrollbar-width: thin;
    }

    .info-nav a {
      flex-shrink: 0;
      border: 1px solid var(--border);
      background: var(--surface);
      border-radius: 999px;
      padding: 9px 16px;
      font-size: 12.5px;
      font-weight: 700;
      color: var(--muted);
      white-space: nowrap;
      box-shadow: var(--shadow);
    }

    .info-nav a:hover {
      border-color: var(--primary);
      color: var(--primary);
    }

    /* ---------- SECTIONS ---------- */
    .section {
      padding: 38px 0;
    }

    .section-head {
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      gap: 16px;
      margin-bottom: 22px;
      flex-wrap: wrap;
    }

    .section-head h3 {
      font-size: 22px;
      margin: 0 0 4px;
      font-weight: 800;
      letter-spacing: -.2px;
    }

    .section-head p {
      margin: 0;
      color: var(--muted);
      font-size: 13.5px;
    }

    .pill-row {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }

    .pill {
      border: 1px solid var(--border);
      background: var(--surface);
      border-radius: 999px;
      padding: 7px 14px;
      font-size: 12.5px;
      font-weight: 600;
      color: var(--muted);
      transition: .12s;
    }

    .pill.active {
      background: var(--primary);
      color: #fff;
      border-color: var(--primary);
    }

    .pill:hover:not(.active) {
      border-color: var(--primary);
      color: var(--primary);
    }

    .subhead {
      font-size: 13px;
      font-weight: 800;
      margin: 26px 0 12px;
      color: var(--text);
      text-transform: uppercase;
      letter-spacing: .4px;
    }

    .subhead:first-of-type {
      margin-top: 4px;
    }

    /* jam layanan */
    .hours-grid {
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      gap: 10px;
    }

    .hour-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 16px 10px;
      text-align: center;
      box-shadow: var(--shadow);
    }

    .hour-card.today {
      border-color: var(--primary);
      background: var(--primary-light);
    }

    .hour-card .day {
      font-weight: 700;
      font-size: 13px;
      margin-bottom: 8px;
    }

    .hour-card .time {
      font-family: var(--font-mono);
      font-size: 12.5px;
      color: var(--text);
    }

    .hour-card .closed {
      color: var(--danger);
      font-weight: 600;
      font-size: 12px;
    }

    .hour-card .note {
      font-size: 10.5px;
      color: var(--muted);
      margin-top: 6px;
    }

    .hour-card.today .day {
      color: var(--primary-dark);
    }

    /* capaian layanan — grafik line berjenjang */
    .charts-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: var(--gap);
    }

    .chart-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-left: 5px solid var(--primary);
      border-radius: var(--radius);
      padding: var(--pad);
      box-shadow: var(--shadow);
      transition: border-color .2s, transform .15s, box-shadow .15s;
      cursor: pointer;
      position: relative;
    }

    .chart-card:hover,
    .chart-card:focus-visible {
      transform: translateY(-3px);
      box-shadow: 0 10px 30px rgba(214, 110, 150, .22);
      outline: none;
    }

    .chart-card:active {
      transform: translateY(-1px);
    }

    .chart-card .click-hint {
      display: flex;
      align-items: center;
      gap: 5px;
      font-size: 10.5px;
      color: var(--primary-dark);
      font-weight: 600;
      margin-top: 10px;
      opacity: .85;
    }

    .chart-card .click-hint svg {
      width: 12px;
      height: 12px;
      flex-shrink: 0;
    }

    /* ---------- MODAL RINCIAN PER DESA (papan informasi publik) ---------- */
    .village-modal {
      max-width: 560px;
    }

    .village-modal .vm-head {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 12px;
    }

    .village-modal .vm-close {
      background: var(--primary-light);
      border: none;
      width: 30px;
      height: 30px;
      border-radius: 50%;
      font-size: 16px;
      color: var(--primary-dark);
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
    }

    .village-modal .vm-sub {
      font-size: 12.5px;
      color: var(--muted);
      margin: 4px 0 0;
    }

    .village-modal .vm-overall {
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: var(--primary-light);
      border-radius: 12px;
      padding: 12px 16px;
      margin: 16px 0 14px;
    }

    .village-modal .vm-overall .lbl {
      font-size: 12px;
      color: var(--primary-dark);
      font-weight: 600;
    }

    .village-modal .vm-overall .val {
      font-family: var(--font-mono);
      font-size: 20px;
      font-weight: 800;
      color: var(--primary-dark);
    }

    .vm-desa-list {
      display: flex;
      flex-direction: column;
      gap: 10px;
    }

    .vm-desa-row {
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 10px 14px;
    }

    .vm-desa-row .vm-desa-top {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      margin-bottom: 6px;
    }

    .vm-desa-row .vm-desa-name {
      font-weight: 700;
      font-size: 13.5px;
    }

    .vm-desa-row .vm-desa-pct {
      font-family: var(--font-mono);
      font-weight: 800;
      font-size: 14.5px;
    }

    .vm-bar-track {
      height: 8px;
      border-radius: 5px;
      background: var(--primary-light);
      overflow: hidden;
    }

    .vm-bar-fill {
      height: 100%;
      border-radius: 5px;
      transition: width .3s;
    }

    .village-modal .vm-foot {
      font-size: 11.5px;
      color: var(--muted);
      margin-top: 16px;
      line-height: 1.5;
    }

    .chart-legend {
      display: flex;
      gap: 14px;
      flex-wrap: wrap;
      font-size: 11.5px;
      color: var(--muted);
      margin: -8px 0 18px;
    }

    .chart-legend span {
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }

    .chart-legend i {
      width: 10px;
      height: 10px;
      border-radius: 3px;
      display: inline-block;
    }

    .chart-card .chead {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 10px;
      margin-bottom: 2px;
    }

    .chart-card h4 {
      margin: 6px 0 0;
      font-size: 14.5px;
      line-height: 1.35;
      font-weight: 700;
    }

    .chart-card .cur {
      text-align: right;
      flex-shrink: 0;
    }

    .chart-card .cur .v {
      font-family: var(--font-mono);
      font-size: 21px;
      font-weight: 700;
      color: var(--primary-dark);
    }

    .chart-card .cur .m {
      font-size: 10.5px;
      color: var(--muted);
      margin-top: 2px;
    }

    .line-chart-svg {
      width: 100%;
      height: auto;
      display: block;
      margin-top: 6px;
    }

    .chart-card .stats {
      font-family: var(--font-mono);
      font-size: 11.5px;
      color: var(--muted);
      margin-top: 6px;
    }

    .chart-card .stats b {
      color: var(--text);
    }

    .program-cat {
      display: inline-block;
      font-size: 10.5px;
      font-weight: 700;
      letter-spacing: .4px;
      text-transform: uppercase;
      color: var(--primary);
      background: var(--primary-light);
      padding: 3px 9px;
      border-radius: 6px;
    }

    /* informasi layanan / fasilitas — kartu sederhana */
    .simple-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: var(--gap);
    }

    .simple-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 18px;
      box-shadow: var(--shadow);
      cursor: pointer;
      transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
    }

    .simple-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 26px rgba(214, 110, 150, 0.2);
      border-color: var(--accent);
    }

    .simple-card h4 {
      margin: 8px 0 6px;
      font-size: 14.5px;
      font-weight: 700;
    }

    .simple-card p {
      margin: 0;
      font-size: 12.5px;
      color: var(--muted);
      line-height: 1.55;
    }

    /* klaster layanan */
    .cluster-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: var(--gap);
    }

    .cluster-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-left: 5px solid var(--primary);
      border-radius: var(--radius);
      padding: 20px;
      box-shadow: var(--shadow);
      cursor: pointer;
      transition: transform .15s ease, box-shadow .15s ease;
    }

    .cluster-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 26px rgba(214, 110, 150, 0.2);
    }

    .cluster-card .num {
      font-family: var(--font-mono);
      font-size: 11px;
      color: var(--primary);
      font-weight: 700;
      letter-spacing: 1px;
      text-transform: uppercase;
    }

    .cluster-card h4 {
      margin: 6px 0 8px;
      font-size: 15.5px;
      font-weight: 800;
    }

    .cluster-card p {
      margin: 0 0 10px;
      font-size: 12.5px;
      color: var(--muted);
      line-height: 1.55;
    }

    .cluster-card ul {
      margin: 0;
      padding-left: 18px;
      font-size: 12.5px;
      color: var(--text);
    }

    .cluster-card li {
      margin-bottom: 4px;
    }

    /* SDM & jadwal dokter */
    .list-stack {
      display: flex;
      flex-direction: column;
      gap: 10px;
      margin-bottom: 8px;
    }

    .doctor-card,
    .staff-card,
    .hotline-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 14px 16px;
      box-shadow: var(--shadow);
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
    }

    .doctor-card,
    .staff-card {
      cursor: pointer;
      transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
    }

    .doctor-card:hover,
    .staff-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 26px rgba(214, 110, 150, 0.2);
      border-color: var(--accent);
    }

    /* isi modal detail (layanan/fasilitas/klaster/SDM/dokter) */
    .vm-body-text {
      font-size: 13.5px;
      color: var(--text);
      line-height: 1.65;
      margin: 0 0 6px;
    }

    .vm-body-list {
      margin: 10px 0 0;
      padding-left: 18px;
      font-size: 13px;
      color: var(--text);
      line-height: 1.6;
    }

    .vm-meta-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 10px 14px;
      margin-top: 10px;
      background: var(--primary-light);
    }

    .vm-meta-row .k {
      font-size: 11.5px;
      color: var(--muted);
      font-weight: 600;
    }

    .vm-meta-row .v {
      font-size: 12.5px;
      color: var(--primary-dark);
      font-weight: 700;
      text-align: right;
    }

    .doctor-card .dname,
    .staff-card .dname,
    .hotline-card .dname {
      font-weight: 700;
      font-size: 13.5px;
    }

    .doctor-card .dspec {
      font-size: 11.5px;
      color: var(--primary);
      font-weight: 600;
      margin-top: 2px;
    }

    .staff-card .dsub,
    .doctor-card .dsub,
    .hotline-card .dsub {
      font-size: 11.5px;
      color: var(--muted);
      margin-top: 2px;
    }

    .doctor-card .dtime,
    .staff-card .dtime,
    .hotline-card .dtime {
      font-family: var(--font-mono);
      font-size: 12px;
      color: var(--muted);
      text-align: right;
      flex-shrink: 0;
    }

    a.hotline-card {
      color: inherit;
      text-decoration: none;
      cursor: pointer;
      transition: transform .12s ease, box-shadow .12s ease, border-color .12s ease;
    }

    a.hotline-card:hover,
    a.hotline-card:focus-visible {
      transform: translateY(-1px);
      border-color: var(--primary);
      box-shadow: 0 6px 20px rgba(214, 110, 150, 0.18);
    }

    a.hotline-card:active {
      transform: translateY(0);
    }

    .hotline-call-btn {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: var(--primary-light);
      color: var(--primary-dark);
      font-weight: 700;
      padding: 7px 12px;
      border-radius: 999px;
      max-width: 280px;
    }

    .hotline-call-btn span {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    @media (max-width: 560px) {
      .hotline-call-btn {
        max-width: 170px;
        font-size: 11px;
      }
    }

    /* breaking news (poster / flyer) */
    .breaking-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
      gap: var(--gap);
    }

    .breaking-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      overflow: hidden;
      box-shadow: var(--shadow);
    }

    .breaking-card .bframe {
      position: relative;
      padding-top: 130%;
      background: #f2e4ea;
    }

    .breaking-card .bframe img {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      object-fit: cover;
    }

    .breaking-card .btitle {
      padding: 12px 14px;
      font-size: 13px;
      font-weight: 700;
    }

    .poster-thumb {
      width: 44px;
      height: 60px;
      object-fit: cover;
      border-radius: 6px;
      border: 1px solid var(--border);
      background: #f2e4ea;
    }

    footer.site-footer {
      background: var(--primary-dark);
      color: #fff;
      padding: 30px 0;
      margin-top: 20px;
    }

    footer.site-footer .container {
      display: flex;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 14px;
      font-size: 12.5px;
      opacity: .85;
    }

    /* ---------- LOGIN ---------- */
    .login-wrap {
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: var(--bg);
      padding: 20px;
    }

    .login-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 34px;
      width: 100%;
      max-width: 380px;
    }

    .login-card .brand-mark {
      background: var(--primary-light);
      color: var(--primary);
      border: 1px solid var(--border);
    }

    .login-card h2 {
      margin: 18px 0 4px;
      font-size: 20px;
      font-weight: 800;
    }

    .login-card p {
      margin: 0 0 22px;
      color: var(--muted);
      font-size: 13px;
    }

    .field {
      margin-bottom: 14px;
    }

    .field label {
      display: block;
      font-size: 12.5px;
      font-weight: 600;
      margin-bottom: 6px;
      color: var(--muted);
    }

    .field input,
    .field select,
    .field textarea {
      width: 100%;
      border: 1px solid var(--border);
      border-radius: 9px;
      padding: 10px 12px;
      font-size: 13.5px;
      background: #fff;
      color: var(--text);
    }

    .field input:focus,
    .field select:focus,
    .field textarea:focus {
      outline: 2px solid var(--primary);
      outline-offset: 1px;
    }

    .field-row {
      display: flex;
      gap: 10px;
    }

    .logo-preview-wrap {
      width: 64px;
      height: 64px;
      border-radius: 50%;
      overflow: hidden;
      flex-shrink: 0;
      background: #fff;
      border: 1px solid var(--border);
      display: flex;
      align-items: center;
      justify-content: center;
    }

    .logo-preview-wrap img {
      width: 100%;
      height: 100%;
      object-fit: contain;
    }

    .poster-preview-wrap {
      width: 90px;
      height: 120px;
      border-radius: 10px;
      overflow: hidden;
      flex-shrink: 0;
      background: #f2e4ea;
      border: 1px solid var(--border);
      display: flex;
      align-items: center;
      justify-content: center;
    }

    .poster-preview-wrap img {
      width: 100%;
      height: 100%;
      object-fit: cover;
    }

    .field-row .field {
      flex: 1;
    }

    .form-msg {
      font-size: 12.5px;
      margin-top: 10px;
      padding: 9px 11px;
      border-radius: 8px;
    }

    .form-msg.err {
      background: #FCEAEA;
      color: var(--danger);
    }

    .form-msg.ok {
      background: var(--primary-light);
      color: var(--primary-dark);
    }

    .demo-note {
      font-size: 11.5px;
      background: #FFF6E8;
      border: 1px solid #F1DDAF;
      color: #8A5A0E;
      padding: 9px 11px;
      border-radius: 8px;
      margin-top: 14px;
    }

    /* ---------- ADMIN LAYOUT ---------- */
    .admin-shell {
      display: flex;
      min-height: 100vh;
    }

    .admin-sidebar {
      width: 250px;
      background: var(--primary-dark);
      color: #fff;
      flex-shrink: 0;
      padding: 20px 14px;
      display: flex;
      flex-direction: column;
      gap: 4px;
      position: sticky;
      top: 0;
      height: 100vh;
      overflow-y: auto;
    }

    .admin-sidebar .brand {
      padding: 6px 8px 20px;
    }

    .admin-sidebar .brand-mark {
      width: 44px;
      height: 44px;
      font-size: 18px;
      border-radius: 50%;
      padding: 3px;
      border: 2px solid var(--border);
      box-shadow: none;
    }

    .admin-sidebar .brand-text h1 {
      font-size: 14px;
    }

    .admin-sidebar .brand-text p {
      font-size: 10.5px;
    }

    .nav-link {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 9px 12px;
      border-radius: 10px;
      color: rgba(255, 255, 255, .78);
      font-size: 13px;
      font-weight: 600;
      background: none;
      border: none;
      text-align: left;
      width: 100%;
    }

    .nav-link:hover {
      background: rgba(255, 255, 255, .08);
      color: #fff;
    }

    .nav-link.active {
      background: rgba(255, 255, 255, .16);
      color: #fff;
    }

    .nav-link .ic {
      width: 18px;
      text-align: center;
    }

    .nav-sep {
      height: 1px;
      background: rgba(255, 255, 255, .12);
      margin: 10px 6px;
    }

    .nav-bottom {
      margin-top: auto;
      padding-top: 10px;
    }

    .admin-main {
      flex: 1;
      min-width: 0;
      padding: 26px 30px 60px;
    }

    .admin-topbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 22px;
      gap: 12px;
      flex-wrap: wrap;
    }

    .admin-topbar h2 {
      margin: 0;
      font-size: 21px;
      font-weight: 800;
    }

    .admin-topbar .sub {
      color: var(--muted);
      font-size: 13px;
      margin-top: 2px;
    }

    .panel {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 22px;
      margin-bottom: 20px;
      box-shadow: var(--shadow);
    }

    .panel h3 {
      margin: 0 0 4px;
      font-size: 15.5px;
      font-weight: 700;
    }

    .panel .desc {
      color: var(--muted);
      font-size: 12.5px;
      margin: 0 0 16px;
    }

    .stat-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 16px;
      margin-bottom: 22px;
    }

    .stat-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 18px;
      box-shadow: var(--shadow);
    }

    .stat-card .k {
      font-family: var(--font-mono);
      font-size: 26px;
      font-weight: 600;
      color: var(--primary-dark);
    }

    .stat-card .l {
      font-size: 12px;
      color: var(--muted);
      margin-top: 4px;
    }

    table.data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }

    table.data-table th {
      text-align: left;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: .4px;
      color: var(--muted);
      border-bottom: 1px solid var(--border);
      padding: 8px 10px;
      font-weight: 700;
    }

    table.data-table td {
      padding: 10px 10px;
      border-bottom: 1px solid var(--border);
      vertical-align: middle;
    }

    table.data-table tr:last-child td {
      border-bottom: none;
    }

    .row-actions {
      display: flex;
      gap: 6px;
    }

    .badge {
      display: inline-block;
      font-size: 10.5px;
      font-weight: 700;
      padding: 3px 9px;
      border-radius: 6px;
      background: var(--primary-light);
      color: var(--primary-dark);
    }

    .badge.off {
      background: #F1F2F0;
      color: var(--muted);
    }

    .badge.role-admin {
      background: var(--primary-light);
      color: var(--primary-dark);
    }

    .badge.role-operator {
      background: #EAF1FB;
      color: #2F5FA8;
    }

    select.pengguna-role-select {
      padding: 7px 10px;
      border: 1px solid var(--border);
      border-radius: 8px;
      font-size: 13px;
      background: var(--surface);
      color: var(--text);
    }

    select.pengguna-role-select:disabled {
      background: #F7F2F4;
      color: var(--muted);
      cursor: not-allowed;
    }

    .toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 10px;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }

    .toolbar .left {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      align-items: center;
    }

    .search-box {
      border: 1px solid var(--border);
      border-radius: 9px;
      padding: 8px 12px;
      font-size: 13px;
      min-width: 200px;
    }

    /* teks berjalan (ticker) — tambah cepat & pratinjau */
    .quick-ticker {
      display: flex;
      gap: 8px;
      margin-bottom: 14px;
    }

    .quick-ticker input {
      flex: 1;
      border: 1px solid var(--border);
      border-radius: 9px;
      padding: 9px 12px;
      font-size: 13px;
    }

    .ticker-preview {
      background: var(--primary-dark);
      color: #fff;
      border-radius: 9px;
      padding: 9px 14px;
      font-size: 12.5px;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
    }

    /* modal */
    .modal-overlay {
      position: fixed;
      inset: 0;
      background: rgba(10, 20, 16, .5);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 100;
      padding: 20px;
    }

    .modal {
      background: #fff;
      border-radius: var(--radius);
      padding: 26px;
      width: 100%;
      max-width: 460px;
      max-height: 90vh;
      overflow-y: auto;
    }

    .modal h3 {
      margin: 0 0 16px;
      font-size: 16.5px;
      font-weight: 800;
    }

    .modal-actions {
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      margin-top: 20px;
    }

    /* toast */
    #toast {
      position: fixed;
      bottom: 22px;
      right: 22px;
      background: var(--primary-dark);
      color: #fff;
      padding: 12px 18px;
      border-radius: 10px;
      font-size: 13px;
      box-shadow: 0 6px 24px rgba(0, 0, 0, .2);
      z-index: 200;
      opacity: 0;
      transform: translateY(8px);
      transition: .2s;
      pointer-events: none;
    }

    #toast.show {
      opacity: 1;
      transform: translateY(0);
    }

    #toast.err {
      background: var(--danger);
    }

    /* color swatches */
    .swatch-row {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }

    .swatch {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 6px;
      font-size: 11px;
      color: var(--muted);
    }

    .swatch input[type=color] {
      width: 44px;
      height: 44px;
      border-radius: 10px;
      border: 1px solid var(--border);
      padding: 2px;
      background: #fff;
    }

    .theme-preset {
      border: 1.5px solid var(--border);
      border-radius: 12px;
      padding: 10px;
      display: flex;
      align-items: center;
      gap: 10px;
      cursor: pointer;
      font-size: 12.5px;
      font-weight: 600;
    }

    .theme-preset .dots {
      display: flex;
      gap: 4px;
    }

    .theme-preset .dots span {
      width: 16px;
      height: 16px;
      border-radius: 50%;
      display: block;
    }

    .theme-preset.selected {
      border-color: var(--primary);
      background: var(--primary-light);
    }

    .preset-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 10px;
    }

    .layout-toggle {
      display: flex;
      gap: 8px;
    }

    .layout-toggle button {
      flex: 1;
    }

    /* pengaturan warna grafik capaian berdasarkan level */
    .chart-color-row {
      display: flex;
      align-items: center;
      gap: 12px;
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 10px 14px;
      margin-bottom: 10px;
      flex-wrap: wrap;
    }

    .chart-color-row .cc-label {
      flex: 1;
      min-width: 140px;
      font-size: 12.5px;
      font-weight: 700;
    }

    .chart-color-row .cc-label small {
      display: block;
      font-weight: 500;
      color: var(--muted);
      font-size: 11px;
      margin-top: 2px;
    }

    .chart-color-row input[type=color] {
      width: 38px;
      height: 38px;
      border-radius: 9px;
      border: 1px solid var(--border);
      padding: 2px;
      background: #fff;
    }

    .chart-color-row input[type=number] {
      width: 70px;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 7px 9px;
      font-size: 12.5px;
    }

    .pf-address-link {
      color: #fff;
      opacity: .85;
      text-decoration: underline;
      font-size: 11.5px;
    }

    .pf-address-link:hover {
      opacity: 1;
    }

    @media(max-width:900px) {
      .charts-grid {
        grid-template-columns: 1fr;
      }

      .simple-grid {
        grid-template-columns: repeat(2, 1fr);
      }

      .cluster-grid {
        grid-template-columns: 1fr;
      }

      .breaking-grid {
        grid-template-columns: repeat(2, 1fr);
      }

      .hours-grid {
        grid-template-columns: repeat(4, 1fr);
      }

      .stat-grid {
        grid-template-columns: repeat(2, 1fr);
      }

      .admin-sidebar {
        position: fixed;
        left: 0;
        top: 0;
        bottom: 0;
        transform: translateX(-100%);
        transition: .2s;
        z-index: 150;
      }

      .admin-sidebar.open {
        transform: translateX(0);
      }

      .admin-main {
        padding: 18px;
      }
    }

    @media(max-width:600px) {
      .simple-grid {
        grid-template-columns: 1fr;
      }

      .hours-grid {
        grid-template-columns: repeat(2, 1fr);
      }

      .hero-headline h2 {
        font-size: 24px;
      }

      .stat-grid {
        grid-template-columns: 1fr 1fr;
      }

      .preset-grid {
        grid-template-columns: 1fr;
      }

      .doctor-card,
      .staff-card {
        flex-direction: column;
        align-items: flex-start;
      }

      .doctor-card .dtime,
      .staff-card .dtime {
        text-align: left;
      }

      .brand-mark {
        width: 58px;
        height: 58px;
      }

      .brand-text h1 {
        font-size: 19px;
      }

      .pemda-logo-inline {
        height: 40px;
      }
    }
  </style>
</head>

<body>

  <!-- =================== TOAST =================== -->
  <div id="toast"></div>

  <!-- =================== PUBLIC VIEW =================== -->
  <div id="view-public" class="hidden">
    <div class="ticker-bar">
      <div id="ticker-content">
        <div class="ticker-empty">Memuat pengumuman…</div>
      </div>
    </div>

    <header class="hero">
      <div class="container">
        <div class="hero-top">
          <div class="brand">
            <img class="pemda-logo-inline"
              src="data:image/png;base64,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"
              alt="Logo Pemerintah Kabupaten Jember">
            <div class="brand-mark"><img id="pf-logo-img" alt="Logo Puskesmas Puger"></div>
            <div class="brand-text">
              <h1 id="pf-name" class="txt-outline">Puskesmas Puger</h1>
              <p id="pf-subtitle">UPT Puskesmas Puger — Dinas Kesehatan Kabupaten Jember</p>
            </div>
          </div>
          <div class="hero-actions">
            <button class="btn btn-ghost" onclick="location.hash='#/login'">🔐 Masuk Admin</button>
          </div>
        </div>
        <div class="hero-headline">
          <span class="eyebrow">📣 Papan Informasi Publik</span>
          <h2 id="pf-tagline" class="txt-outline-lg">Melayani <span class="hl-word">Sepenuh Hati</span> untuk Masyarakat
            Puger</h2>
          <p>Pantau jam layanan, capaian program, dan seluruh informasi kesehatan Puskesmas Puger secara terbuka dan
            diperbarui berkala.</p>
        </div>
      </div>
      <svg class="wave-divider" viewBox="0 0 1440 80" preserveAspectRatio="none">
        <path d="M0,40 C240,90 480,0 720,30 C960,60 1200,10 1440,45 L1440,80 L0,80 Z"></path>
      </svg>
    </header>

    <main class="container">

      <!-- MENU PENGELOMPOKAN INFORMASI PUBLIK -->
      <nav class="info-nav">
        <a href="#sec-hours">🕒 Jam Layanan</a>
        <a href="#sec-capaian">📈 Capaian Layanan</a>
        <a href="#sec-layanan">🩺 Informasi Jenis Layanan</a>
        <a href="#sec-fasilitas">🏥 Fasilitas Kesehatan</a>
        <a href="#sec-klaster">🧩 Klaster Layanan</a>
        <a href="#sec-sdm">👩‍⚕️ SDM &amp; Jadwal Dokter</a>
        <a href="#sec-hotline">📞 Hotline &amp; Kontak</a>
        <a href="#sec-breaking">📰 Breaking News</a>
      </nav>

      <section class="section" id="sec-hours">
        <div class="section-head">
          <div>
            <h3>Jam Layanan</h3>
            <p>Jadwal pelayanan mingguan Puskesmas Puger</p>
          </div>
        </div>
        <div class="hours-grid" id="hours-grid"></div>
      </section>

      <section class="section" id="sec-capaian">
        <div class="section-head">
          <div>
            <h3>Capaian Layanan</h3>
            <p>Perkembangan capaian program kesehatan berjenjang dari Januari sampai Desember</p>
          </div>
          <div class="pill-row" id="category-filter"></div>
        </div>
        <div class="chart-legend" id="capaian-legend"></div>
        <div class="charts-grid" id="capaian-grid"></div>
      </section>

      <section class="section" id="sec-layanan">
        <div class="section-head">
          <div>
            <h3>Informasi Jenis Layanan</h3>
            <p>Jenis-jenis layanan kesehatan yang tersedia di Puskesmas Puger</p>
          </div>
        </div>
        <div class="simple-grid" id="layanan-grid"></div>
      </section>

      <section class="section" id="sec-fasilitas">
        <div class="section-head">
          <div>
            <h3>Fasilitas Layanan Kesehatan</h3>
            <p>Sarana dan prasarana pendukung pelayanan</p>
          </div>
        </div>
        <div class="simple-grid" id="fasilitas-grid"></div>
      </section>

      <section class="section" id="sec-klaster">
        <div class="section-head">
          <div>
            <h3>Klaster Layanan</h3>
            <p>Pengelompokan layanan berbasis Integrasi Layanan Primer (ILP), Klaster 1–4</p>
          </div>
        </div>
        <div class="cluster-grid" id="cluster-grid"></div>
      </section>

      <section class="section" id="sec-sdm">
        <div class="section-head">
          <div>
            <h3>Informasi SDM &amp; Jadwal Dokter</h3>
            <p>Kebutuhan tenaga kesehatan dan jadwal kunjungan dokter spesialis</p>
          </div>
        </div>
        <div class="subhead">Pengumuman Kebutuhan Tenaga Kesehatan dan Tenaga Administrasi</div>
        <div class="list-stack" id="staff-list"></div>
        <button type="button" class="btn-outline" style="margin-top:10px;" onclick="openApplicationUploadModal()">📤
          Upload Berkas Pendaftaran</button>
        <div class="subhead">Jadwal Kunjungan Dokter Spesialis</div>
        <div class="list-stack" id="doctor-list"></div>
      </section>

      <section class="section" id="sec-hotline">
        <div class="section-head">
          <div>
            <h3>Hotline &amp; Kontak</h3>
            <p>Nomor layanan call inbound dan kontak darurat Puskesmas Puger</p>
          </div>
        </div>
        <div class="list-stack" id="hotline-list"></div>
      </section>

      <section class="section" id="sec-breaking">
        <div class="section-head">
          <div>
            <h3>Breaking News</h3>
            <p>Poster dan flyer informasi terbaru Puskesmas Puger</p>
          </div>
        </div>
        <div class="breaking-grid" id="breaking-grid"></div>
      </section>

    </main>

    <footer class="site-footer">
      <div class="container">
        <div>
          <div id="pf-address">Jl. A. Yani Nomor 32, Pugerkulon , Kecamatan Puger, Kabupaten Jember</div>
          <div id="pf-phone" style="margin-top:4px;opacity:.8;"></div>
          <a id="pf-maps" class="pf-address-link hidden" href="#" target="_blank" rel="noopener"
            style="display:inline-block;margin-top:6px;">📍 Lihat di Google Maps</a>
        </div>
        <div id="pf-emergency"></div>
      </div>
    </footer>
  </div>

  <!-- =================== LOGIN VIEW =================== -->
  <div id="view-login" class="hidden">
    <div class="login-wrap">
      <div class="login-card">
        <div class="brand-mark">🔐</div>
        <h2>Masuk sebagai Admin</h2>
        <p>Kelola papan informasi &amp; capaian program Puskesmas Puger.</p>
        <form id="login-form">
          <div class="field">
            <label>Alamat Email</label>
            <input type="email" id="login-email" required placeholder="puskesmaspuger123@gmail.com">
          </div>
          <div class="field">
            <label>Kata Sandi</label>
            <input type="password" id="login-password" required placeholder="******">
          </div>
          <button type="submit" class="btn btn-primary" style="width:100%;justify-content:center;">Masuk</button>
        </form>
        <div id="login-msg"></div>
        <div id="demo-note" class="demo-note hidden">
          Supabase belum terhubung. Isi <code>SUPABASE_URL</code> dan <code>SUPABASE_ANON_KEY</code> pada bagian
          CONFIG di berkas ini agar login admin dapat berfungsi.
        </div>
        <div style="text-align:center;margin-top:18px;">
          <a href="#/" style="font-size:12.5px;color:var(--muted);">← Kembali ke papan informasi</a>
        </div>
      </div>
    </div>
  </div>

  <!-- =================== ADMIN VIEW =================== -->
  <div id="view-admin" class="hidden">
    <div class="admin-shell">
      <aside class="admin-sidebar" id="admin-sidebar">
        <div class="brand">
          <div class="brand-mark"><img id="admin-logo-img" alt="Logo Puskesmas Puger"></div>
          <div class="brand-text">
            <h1>Puskesmas Puger</h1>
            <p>Panel Admin</p>
          </div>
        </div>
        <button class="nav-link" data-tab="ringkasan"><span class="ic">📊</span> Ringkasan</button>
        <button class="nav-link" data-tab="program"><span class="ic">📈</span> Capaian Layanan</button>
        <button class="nav-link" data-tab="pengumuman"><span class="ic">📣</span> Papan Informasi</button>
        <button class="nav-link" data-tab="layanan"><span class="ic">🩺</span> Informasi Jenis Layanan</button>
        <button class="nav-link" data-tab="fasilitas"><span class="ic">🏥</span> Fasilitas Kesehatan</button>
        <button class="nav-link" data-tab="klaster"><span class="ic">🧩</span> Klaster Layanan</button>
        <button class="nav-link" data-tab="sdm"><span class="ic">👩‍⚕️</span> SDM &amp; Jadwal Dokter</button>
        <button class="nav-link" data-tab="hotline"><span class="ic">📞</span> Hotline &amp; Kontak</button>
        <button class="nav-link" data-tab="breaking"><span class="ic">📰</span> Breaking News</button>
        <button class="nav-link" data-tab="jam"><span class="ic">🕒</span> Jam Layanan</button>
        <button class="nav-link" data-tab="tampilan"><span class="ic">🎨</span> Pengaturan Tampilan</button>
        <button class="nav-link" data-tab="akun"><span class="ic">👤</span> Akun</button>
        <div class="nav-sep"></div>
        <button class="nav-link" onclick="location.hash='#/'"><span class="ic">🌐</span> Lihat Papan Publik</button>
        <div class="nav-bottom">
          <button class="nav-link" id="btn-logout"><span class="ic">↩</span> Keluar</button>
        </div>
      </aside>

      <main class="admin-main">
        <div class="admin-topbar">
          <div>
            <h2 id="admin-title">Ringkasan</h2>
            <div class="sub" id="admin-sub">Ikhtisar data papan informasi</div>
          </div>
          <button class="btn btn-outline hidden" id="btn-menu-mobile"
            onclick="document.getElementById('admin-sidebar').classList.toggle('open')">☰ Menu</button>
        </div>

        <!-- RINGKASAN -->
        <div class="admin-tab" id="tab-ringkasan">
          <div class="stat-grid">
            <div class="stat-card">
              <div class="k" id="stat-programs">0</div>
              <div class="l">Program Tercatat</div>
            </div>
            <div class="stat-card">
              <div class="k" id="stat-avg">0%</div>
              <div class="l">Rata-rata Capaian Terbaru</div>
            </div>
            <div class="stat-card">
              <div class="k" id="stat-announce">0</div>
              <div class="l">Teks Berjalan Aktif</div>
            </div>
            <div class="stat-card">
              <div class="k" id="stat-days">0/7</div>
              <div class="l">Hari Buka Layanan</div>
            </div>
          </div>
          <div class="panel">
            <h3>Capaian Tertinggi &amp; Terendah</h3>
            <p class="desc">Ringkasan cepat performa program berjalan (berdasarkan bulan terakhir yang terisi)</p>
            <div id="ringkasan-list"></div>
          </div>
        </div>

        <!-- PROGRAM & CAPAIAN -->
        <div class="admin-tab hidden" id="tab-program">
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <input class="search-box" id="program-search" placeholder="Cari program…">
                <select class="search-box" id="program-cat-filter">
                  <option value="">Semua Kategori</option>
                </select>
              </div>
              <div class="left">
                <button class="btn btn-outline btn-sm" id="btn-export-xlsx">⬇ Ekspor Excel (.xlsx)</button>
                <button class="btn btn-outline btn-sm" id="btn-export-csv">⬇ Ekspor CSV</button>
                <label class="btn btn-outline btn-sm" style="margin:0;">⬆ Impor Excel/CSV
                  <input type="file" id="csv-input" accept=".xlsx,.xls,.csv" class="hidden">
                </label>
                <button class="btn btn-primary btn-sm" id="btn-add-program">+ Tambah Data Bulan</button>
              </div>
            </div>
            <p class="desc" style="margin-top:-6px;">Format kolom mengikuti template Excel resmi (<code>Report.xlsx</code>):
              <code>name, category, desa, period year, month, target, achieved, unit</code> (boleh juga file .csv dengan
              kolom yang sama; header <code>period</code> tanpa "year" tetap dikenali). Kolom <code>month</code> bisa diisi
              angka 1–12 atau nama bulan (mis. Januari/Jan). Kolom <code>desa</code> diisi <code>Semua Desa</code>
              untuk data total Puskesmas (dipakai grafik utama), atau nama salah satu desa binaan untuk mengisi
              rincian per desa; boleh dikosongkan dan otomatis dianggap Semua Desa — <b>desa lain yang belum diisi
              akan otomatis diperkirakan sistem</b> di halaman publik, jadi tidak wajib mengisi tiap desa satu per
              satu. Satu baris = data satu program/desa pada satu bulan, sehingga capaian bisa ditampilkan
              berjenjang Januari–Desember pada grafik line di halaman publik. Gunakan tombol <b>Ekspor Excel</b>
              untuk mendapatkan file dengan format dan urutan kolom yang persis sama dengan template, supaya saat
              diimpor kembali datanya tidak berantakan.
            </p>
            <div style="overflow-x:auto;">
              <table class="data-table">
                <thead>
                  <tr>
                    <th>Program</th>
                    <th>Kategori</th>
                    <th>Desa</th>
                    <th>Periode</th>
                    <th>Bulan</th>
                    <th>Target</th>
                    <th>Capaian</th>
                    <th>%</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody id="program-tbody"></tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- PENGUMUMAN / TEKS BERJALAN -->
        <div class="admin-tab hidden" id="tab-pengumuman">
          <div class="panel">
            <h3>Tambah Cepat Teks Berjalan</h3>
            <p class="desc">Ketik teks singkat lalu klik Tambahkan — langsung tampil di teks berjalan (ticker) halaman
              publik.</p>
            <div class="quick-ticker">
              <input id="quick-ticker-input" placeholder="Contoh: Vaksinasi booster gratis setiap Senin–Jumat">
              <button class="btn btn-primary btn-sm" id="btn-quick-ticker">+ Tambahkan</button>
            </div>
            <div class="ticker-preview" id="ticker-live-preview">Pratinjau teks berjalan akan tampil di sini…</div>
          </div>
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Kelola seluruh teks berjalan — ubah urutan prioritas, edit, atau
                  nonaktifkan</p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-announce">+ Tambah Pengumuman</button>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Judul</th>
                  <th>Prioritas</th>
                  <th>Status</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="announce-tbody"></tbody>
            </table>
          </div>
        </div>

        <!-- INFORMASI JENIS LAYANAN -->
        <div class="admin-tab hidden" id="tab-layanan">
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Kelola informasi jenis layanan yang tersedia di Puskesmas Puger</p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-service">+ Tambah Layanan</button>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Nama</th>
                  <th>Kategori</th>
                  <th>Deskripsi</th>
                  <th>Status</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="service-tbody"></tbody>
            </table>
          </div>
        </div>

        <!-- FASILITAS KESEHATAN -->
        <div class="admin-tab hidden" id="tab-fasilitas">
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Kelola daftar fasilitas layanan kesehatan</p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-facility">+ Tambah Fasilitas</button>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Nama</th>
                  <th>Kategori</th>
                  <th>Deskripsi</th>
                  <th>Status</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="facility-tbody"></tbody>
            </table>
          </div>
        </div>

        <!-- KLASTER LAYANAN -->
        <div class="admin-tab hidden" id="tab-klaster">
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <h3 style="margin:0 0 4px;">Klaster Layanan (Integrasi Layanan Primer)</h3>
                <p class="desc" style="margin:0;">Struktur klaster umumnya mengikuti pedoman ILP (Klaster 1–4), namun
                  Anda dapat menambah, mengubah, atau menghapus klaster sesuai kebutuhan.</p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-cluster">+ Tambah Klaster</button>
            </div>
          </div>
          <div id="cluster-admin-list"></div>
        </div>

        <!-- SDM & JADWAL DOKTER -->
        <div class="admin-tab hidden" id="tab-sdm">
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Kelola pengumuman kebutuhan tenaga kesehatan dan tenaga administrasi
                </p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-staff">+ Tambah Kebutuhan</button>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Posisi</th>
                  <th>Kualifikasi</th>
                  <th>Jumlah</th>
                  <th>Batas Lamar</th>
                  <th>Formulir</th>
                  <th>Status</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="staff-tbody"></tbody>
            </table>
          </div>
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Kelola jadwal kunjungan dokter spesialis</p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-doctor">+ Tambah Jadwal Dokter</button>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Nama Dokter</th>
                  <th>Spesialisasi</th>
                  <th>Hari</th>
                  <th>Jam</th>
                  <th>Lokasi</th>
                  <th>Status</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="doctor-tbody"></tbody>
            </table>
          </div>

          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Kode verifikasi Google Authenticator — pendaftar harus memasukkan kode
                  ini saat mengunggah berkas lamaran</p>
              </div>
            </div>
            <div id="totp-panel"></div>
          </div>

          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Berkas pendaftaran yang masuk dari pelamar</p>
              </div>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Waktu</th>
                  <th>Posisi</th>
                  <th>Nama</th>
                  <th>Kontak</th>
                  <th>Berkas</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="applications-tbody"></tbody>
            </table>
          </div>
        </div>

        <!-- Hotline & Kontak -->
        <div class="admin-tab hidden" id="tab-hotline">
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Kelola hotline dan informasi kontak</p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-hotline">+ Tambah Hotline</button>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Layanan Call Inbound</th>
                  <th>Nomor / Link Hotline</th>
                  <th>Status</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="hotline-tbody"></tbody>
            </table>
          </div>
        </div>

        <!-- BREAKING NEWS (POSTER / FLYER) -->
        <div class="admin-tab hidden" id="tab-breaking">
          <div class="panel">
            <div class="toolbar">
              <div class="left">
                <p class="desc" style="margin:0;">Unggah gambar poster atau flyer. Tidak dibatasi satu — tambahkan
                  sebanyak yang diperlukan.</p>
              </div>
              <button class="btn btn-primary btn-sm" id="btn-add-breaking">+ Tambah Poster/Flyer</button>
            </div>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Poster</th>
                  <th>Judul</th>
                  <th>Status</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="breaking-tbody"></tbody>
            </table>
          </div>
        </div>

        <!-- JAM LAYANAN -->
        <div class="admin-tab hidden" id="tab-jam">
          <div class="panel">
            <h3>Jam Layanan per Hari</h3>
            <p class="desc">Atur jam buka, tutup, dan status libur untuk tiap hari. Perubahan langsung tampil di papan
              publik setelah disimpan.</p>
            <table class="data-table">
              <thead>
                <tr>
                  <th>Hari</th>
                  <th>Status</th>
                  <th>Jam Buka</th>
                  <th>Jam Tutup</th>
                  <th>Catatan</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="hours-tbody"></tbody>
            </table>
          </div>
        </div>

        <!-- PENGATURAN TAMPILAN -->
        <div class="admin-tab hidden" id="tab-tampilan">
          <div class="panel">
            <h3>Profil Puskesmas</h3>
            <p class="desc">Informasi ini tampil di header dan footer papan publik</p>
            <div class="field">
              <label>Logo Profil Puskesmas</label>
              <div style="display:flex;align-items:center;gap:14px;">
                <div class="logo-preview-wrap"><img id="set-logo-preview" alt="Pratinjau logo profil"></div>
                <div>
                  <input type="file" id="set-logo-file" accept="image/png,image/jpeg,image/webp" class="hidden">
                  <button type="button" class="btn btn-outline btn-sm" id="btn-pick-logo">Pilih Gambar</button>
                  <button type="button" class="btn btn-outline btn-sm" id="btn-remove-logo">Gunakan Default</button>
                  <p class="desc" style="margin:6px 0 0;">Logo tampil di header papan publik & sidebar admin. Ukuran
                    otomatis disesuaikan. Klik "Simpan Profil" untuk menerapkan.</p>
                </div>
              </div>
            </div>
            <div class="field-row">
              <div class="field"><label>Nama Puskesmas</label><input id="set-name"></div>
              <div class="field"><label>Sub-judul</label><input id="set-subtitle"></div>
            </div>
            <div class="field"><label>Tagline</label><input id="set-tagline"></div>
            <div class="field-row">
              <div class="field"><label>Alamat</label><input id="set-address"></div>
              <div class="field"><label>Telepon</label><input id="set-phone"></div>
            </div>
            <div class="field"><label>Info Layanan Darurat</label><input id="set-emergency"></div>
            <div class="field"><label>Link Google Maps (opsional)</label><input id="set-maps"
                placeholder="https://maps.app.goo.gl/..."></div>
            <button class="btn btn-primary btn-sm" id="btn-save-profile">Simpan Profil</button>
          </div>

          <div class="panel">
            <h3>Warna &amp; Tema</h3>
            <p class="desc">Pilih preset atau atur warna kustom untuk papan informasi</p>
            <div class="preset-grid" id="preset-grid"></div>
            <div style="height:16px;"></div>
            <div class="swatch-row">
              <div class="swatch"><input type="color" id="col-primary" value="#EC9CB6">Utama</div>
              <div class="swatch"><input type="color" id="col-primary-dark" value="#C26A88">Utama Gelap</div>
              <div class="swatch"><input type="color" id="col-accent" value="#F4B8CE">Aksen</div>
              <div class="swatch"><input type="color" id="col-bg" value="#FFF6F9">Latar</div>
            </div>
            <div style="height:20px;"></div>
            <h3 style="font-size:13.5px;">Kepadatan Tata Letak</h3>
            <div class="layout-toggle">
              <button class="btn btn-outline btn-sm" data-layout="comfortable">Nyaman (lega)</button>
              <button class="btn btn-outline btn-sm" data-layout="compact">Ringkas (padat)</button>
            </div>
            <div style="height:20px;"></div>
            <h3 style="font-size:13.5px;">Kecepatan Teks Berjalan</h3>
            <p class="desc">Geser untuk mengatur cepat/lambat teks berjalan (ticker) di papan publik</p>
            <div style="display:flex;align-items:center;gap:12px;">
              <span style="font-size:11.5px;color:var(--muted);">Cepat</span>
              <input type="range" id="ticker-speed" min="10" max="60" step="2" style="flex:1;">
              <span style="font-size:11.5px;color:var(--muted);">Lambat</span>
            </div>
            <div style="height:20px;"></div>
            <button class="btn btn-primary btn-sm" id="btn-save-theme">Simpan Tampilan</button>
            <button class="btn btn-outline btn-sm" id="btn-reset-theme">Kembalikan Default</button>
          </div>

          <div class="panel">
            <h3>Warna Grafik Capaian Program</h3>
            <p class="desc">Atur agar warna garis grafik capaian berubah otomatis sesuai level pencapaian, atau gunakan
              satu warna tetap.</p>
            <div class="layout-toggle" style="margin-bottom:16px;">
              <button class="btn btn-outline btn-sm" data-chartmode="auto">Otomatis sesuai capaian</button>
              <button class="btn btn-outline btn-sm" data-chartmode="fixed">Warna tetap</button>
            </div>
            <div id="chart-color-auto-fields">
              <div class="chart-color-row">
                <div class="cc-label">Capaian Baik<small>≥ ambang batas di samping (%)</small></div>
                <input type="number" id="cc-threshold-excellent" min="0" max="200">
                <input type="color" id="cc-color-excellent">
              </div>
              <div class="chart-color-row">
                <div class="cc-label">Capaian Cukup<small>≥ ambang batas di samping (%)</small></div>
                <input type="number" id="cc-threshold-good" min="0" max="200">
                <input type="color" id="cc-color-good">
              </div>
              <div class="chart-color-row">
                <div class="cc-label">Capaian Kurang<small>di bawah ambang batas "Cukup"</small></div>
                <input type="color" id="cc-color-poor">
              </div>
            </div>
            <div id="chart-color-fixed-fields" class="hidden">
              <div class="chart-color-row">
                <div class="cc-label">Warna Garis Grafik<small>Digunakan untuk semua grafik capaian</small></div>
                <input type="color" id="cc-color-fixed">
              </div>
            </div>
            <button class="btn btn-primary btn-sm" id="btn-save-chartcolor">Simpan Warna Grafik</button>
          </div>
        </div>

        <!-- AKUN -->
        <div class="admin-tab hidden" id="tab-akun">
          <div class="panel">
            <h3>Akun Admin</h3>
            <p class="desc">Kelola email dan kata sandi yang digunakan untuk masuk ke panel admin ini.</p>
            <div id="akun-demo-note" class="form-msg hidden" style="background:#FFF6E5;color:#8A5A00;">
              Edit akun hanya aktif setelah aplikasi terhubung ke Supabase. Isi <code>SUPABASE_URL</code> dan
              <code>SUPABASE_ANON_KEY</code> pada bagian CONFIG di berkas ini terlebih dahulu.
            </div>
            <div class="field"><label>Email Saat Ini</label><input id="akun-current-email" disabled></div>
            <div class="field"><label>Email Baru <small style="font-weight:400;color:var(--muted);">(kosongkan jika
                  tidak diubah)</small></label><input id="akun-new-email" type="email"
                placeholder="puskesmaspuger123@gmail.com"></div>
            <div class="field-row">
              <div class="field"><label>Password Baru <small style="font-weight:400;color:var(--muted);">(kosongkan jika
                    tidak diubah)</small></label><input id="akun-new-password" type="password"
                  placeholder="Minimal 6 karakter"></div>
              <div class="field"><label>Konfirmasi Password Baru</label><input id="akun-confirm-password"
                  type="password" placeholder="Ulangi password baru"></div>
            </div>
            <button class="btn btn-primary btn-sm" id="btn-save-akun">Simpan Akun</button>
            <div id="akun-msg"></div>
          </div>

          <div class="panel">
            <h3>Kelola Pengguna &amp; Peran</h3>
            <p class="desc">Atur peran (role) setiap akun admin dan nonaktifkan akun yang tidak lagi digunakan.
              Akun berperan <b>Operator</b> hanya dapat mengelola konten, sedangkan <b>Admin</b> juga dapat
              mengelola akun lain di menu ini. Akun yang dinonaktifkan akan langsung ditolak saat mencoba masuk.</p>
            <div id="pengguna-demo-note" class="form-msg hidden" style="background:#FFF6E5;color:#8A5A00;">
              Fitur ini hanya aktif setelah aplikasi terhubung ke Supabase dan tabel <code>admin_profiles</code>
              sudah dibuat (lihat komentar CONFIG pada awal berkas ini).
            </div>
            <div id="pengguna-msg"></div>
            <div class="field" style="max-width:320px;">
              <input id="pengguna-search" type="search" placeholder="Cari berdasarkan email...">
            </div>
            <div style="overflow-x:auto;">
              <table class="data-table" id="pengguna-table">
                <thead>
                  <tr>
                    <th>Email</th>
                    <th>Peran</th>
                    <th>Status</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody id="pengguna-tbody"></tbody>
              </table>
            </div>
          </div>

          <div class="panel" id="audit-log-panel">
            <h3>Riwayat Aktivitas Akun</h3>
            <p class="desc">Catatan siapa mengubah peran, menonaktifkan/mengaktifkan, menambah akun, atau meminta
              reset password akun lain. Hanya terlihat oleh Admin.</p>
            <div id="audit-log-demo-note" class="form-msg hidden" style="background:#FFF6E5;color:#8A5A00;">
              Riwayat ini aktif setelah tabel <code>admin_audit_log</code> dibuat di Supabase (lihat komentar
              CONFIG pada awal berkas ini). Tanpa tabel ini, aksi tetap berjalan normal, hanya riwayatnya tidak
              tercatat.
            </div>
            <div style="overflow-x:auto;">
              <table class="data-table" id="audit-log-table">
                <thead>
                  <tr>
                    <th>Waktu</th>
                    <th>Oleh</th>
                    <th>Aksi</th>
                    <th>Target</th>
                    <th>Detail</th>
                  </tr>
                </thead>
                <tbody id="audit-log-tbody"></tbody>
              </table>
            </div>
          </div>

          <div class="panel" id="tambah-akun-panel">
            <h3>Tambah Akun Admin Baru</h3>
            <p class="desc">Buat akun admin tambahan agar lebih dari satu orang bisa masuk ke panel ini. Pilih
              peran akun baru di bawah ini — Anda selalu dapat mengubahnya kembali lewat tabel "Kelola Pengguna
              &amp; Peran" di atas.</p>
            <div id="tambah-akun-demo-note" class="form-msg hidden" style="background:#FFF6E5;color:#8A5A00;">
              Fitur ini hanya aktif setelah aplikasi terhubung ke Supabase.
            </div>
            <div class="field-row">
              <div class="field"><label>Email Akun Baru</label><input id="ta-email" type="email"
                  placeholder="nama@contoh.com"></div>
              <div class="field"><label>Password Akun Baru</label><input id="ta-password" type="password"
                  placeholder="Minimal 6 karakter"></div>
            </div>
            <div class="field-row">
              <div class="field"><label>Konfirmasi Password</label><input id="ta-confirm-password" type="password"
                  placeholder="Ulangi password"></div>
              <div class="field"><label>Peran</label>
                <select id="ta-role">
                  <option value="admin">Admin (akses penuh, termasuk kelola pengguna)</option>
                  <option value="operator">Operator (kelola konten saja)</option>
                </select>
              </div>
            </div>
            <button class="btn btn-primary btn-sm" id="btn-tambah-akun">Tambah Akun</button>
            <div id="tambah-akun-msg"></div>
            <p class="desc" style="margin-top:10px;">Catatan: jika opsi "Confirm email" masih aktif di pengaturan
              Authentication Supabase Anda, akun baru harus mengklik tautan konfirmasi yang dikirim ke emailnya
              sebelum bisa masuk. Untuk mengizinkan akun baru langsung masuk tanpa konfirmasi email, matikan opsi
              tersebut di Supabase Dashboard &gt; Authentication &gt; Providers &gt; Email &gt; "Confirm email", atau
              konfirmasi akun secara manual lewat Authentication &gt; Users.</p>
          </div>
        </div>
      </main>
    </div>
  </div>

  <!-- =================== MODAL (generik, diisi lewat JS) =================== -->
  <div id="modal-root"></div>

  <script>
    /* =====================================================================
       LOGO PROFIL DEFAULT (fallback) — dipakai di header publik & sidebar
       admin selama admin belum mengunggah logo sendiri lewat menu
       Pengaturan Tampilan > Profil Puskesmas. Disimpan sekali di sini
       (bukan diulang di beberapa tag <img>) agar ukuran berkas tetap hemat.
       ===================================================================== */
    const DEFAULT_PROFILE_LOGO = "data:image/png;base64,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";
    (function applyDefaultLogoImmediately() {
      const a = document.getElementById('pf-logo-img'); if (a) a.src = DEFAULT_PROFILE_LOGO;
      const b = document.getElementById('admin-logo-img'); if (b) b.src = DEFAULT_PROFILE_LOGO;
    })();

    /* =====================================================================
       CONFIG — isi dengan Project URL & anon public key dari Supabase Anda
       (Dashboard Supabase > Settings > API > Project URL & anon public key).

       WAJIB diisi agar aplikasi ini menjadi website Supabase ASLI (bukan
       mode demo). Sebelum mengisi, jalankan skrip "supabase-schema.sql"
       yang disertakan di SQL Editor proyek Supabase Anda, lalu buat akun
       admin di Authentication > Users dengan:
         Email    : puskesmaspuger123@gmail.com
         Password : puger12345

       Selama SUPABASE_URL / SUPABASE_ANON_KEY masih kosong, aplikasi tidak
       akan bisa login sama sekali (akun demo sudah dihapus dari aplikasi ini).

       UNTUK FITUR "Role & Nonaktifkan Akun" (menu Pengaturan > Akun):
       jalankan juga skrip tabel "admin_profiles" berikut ini satu kali di
       SQL Editor Supabase (boleh disatukan dengan supabase-schema.sql):

         create table if not exists public.admin_profiles (
           id uuid primary key references auth.users(id) on delete cascade,
           email text not null,
           role text not null default 'admin' check (role in ('admin','operator')),
           is_active boolean not null default true,
           created_at timestamptz not null default now()
         );
         alter table public.admin_profiles enable row level security;
         create policy "admin_profiles_select" on public.admin_profiles
           for select using (auth.role() = 'authenticated');
         create policy "admin_profiles_insert" on public.admin_profiles
           for insert with check (auth.role() = 'authenticated');
         create policy "admin_profiles_update" on public.admin_profiles
           for update using (auth.role() = 'authenticated');

       Tabel ini dipakai untuk menyimpan peran (role) dan status aktif/nonaktif
       tiap akun admin. Supabase tidak mengizinkan menonaktifkan/menghapus akun
       auth.users langsung dari sisi browser (perlu service_role key yang tidak
       aman disimpan di file ini), sehingga "nonaktif" di sini berarti: begitu
       akun tersebut login, aplikasi otomatis akan menolak & mengeluarkannya.

       UNTUK FITUR "Riwayat Aktivitas Akun" (log siapa mengubah peran/menonaktifkan
       siapa, di menu Pengaturan > Akun): jalankan juga skrip tabel berikut ini
       satu kali di SQL Editor Supabase (boleh disatukan dengan skrip di atas):

         create table if not exists public.admin_audit_log (
           id uuid primary key default gen_random_uuid(),
           actor_id uuid,
           actor_email text,
           action text not null,
           target_email text,
           detail text,
           created_at timestamptz not null default now()
         );
         alter table public.admin_audit_log enable row level security;
         create policy "admin_audit_log_select" on public.admin_audit_log
           for select using (auth.role() = 'authenticated');
         create policy "admin_audit_log_insert" on public.admin_audit_log
           for insert with check (auth.role() = 'authenticated');

       Tabel ini hanya mencatat riwayat (siapa mengubah apa dan kapan) dan tidak
       memengaruhi jalannya aplikasi bila belum dibuat — bagian "Riwayat Aktivitas
       Akun" akan menampilkan pesan agar tabel ini dibuat terlebih dahulu.
       ===================================================================== */
    const SUPABASE_URL = "https://ptalsoqynzlsyuarltvs.supabase.co";
    const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB0YWxzb3F5bnpsc3l1YXJsdHZzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODY4MDQ3MzIsImV4cCI6MjEwMjM4MDczMn0.OAhTVVYZhyI4ZRi28hdsHmLD4GCV0VUtWwB6kqw1AnQ"
    const DEMO_MODE = !SUPABASE_URL || !SUPABASE_ANON_KEY;
    const sb = DEMO_MODE ? null : supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    /* Klien Supabase KEDUA, khusus dipakai untuk mendaftarkan akun admin baru (menu Akun >
       Tambah Akun Admin Baru). Wajib terpisah dari klien "sb" di atas: signUp() pada klien yang
       sama akan mengganti sesi login yang sedang aktif dengan sesi akun baru, sehingga admin yang
       sedang login otomatis "terganti" oleh akun yang baru saja dia buat. Dengan storageKey unik
       dan persistSession:false, klien ini tidak menyentuh sesi admin yang sedang login. */
    const sbForNewAdmin = DEMO_MODE ? null : supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
      auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false, storageKey: 'sb-tambah-akun-temp' }
    });

    /* Ambil pesan error yang aman ditampilkan ke pengguna. Beberapa error (mis. dari
       Supabase, atau event error murni dari browser) bukan instance Error dan tidak
       punya properti .message, sehingga jika ditempel langsung ke string akan tampil
       sebagai "[object Object]". Fungsi ini mengekstrak pesan yang paling masuk akal
       dari berbagai bentuk error tersebut. */
    function errMsg(e) {
      if (!e) return 'Terjadi kesalahan yang tidak diketahui.';
      if (typeof e === 'string') return e;
      if (e.message) return e.message;
      if (e.error_description) return e.error_description;
      if (e.error) return typeof e.error === 'string' ? e.error : errMsg(e.error);
      if (e.msg) return e.msg;
      try {
        const s = JSON.stringify(e);
        if (s && s !== '{}') return s;
      } catch (_) { /* abaikan, misal ada referensi siklik */ }
      return String(e);
    }

    const DAY_NAMES = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
    const MONTH_NAMES = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const MONTH_SHORT = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];

    /* Desa/kelurahan binaan wilayah kerja Puskesmas Puger — dipilih admin lewat
       menu Desa pada form edit "Capaian Layanan" agar rincian capaian per desa
       yang tampil saat kartu diklik publik berasal dari data asli (lihat
       villageBreakdownFor() dan DESA_ALL). */
    const DESA_LIST = ['Pugerkulon', 'Pugerwetan', 'Mojosari', 'Mojomulyo', 'Grenden'];
    /* Nilai desa untuk baris data agregat/total Puskesmas (bukan data satu desa tertentu).
       Baris lama yang belum punya kolom desa dianggap otomatis sebagai baris agregat ini. */
    const DESA_ALL = 'Semua Desa';
    function isAggregateDesa(d) { return !d || d === DESA_ALL; }

    function parseMonth(v) {
      if (v === undefined || v === null || v === '') return 1;
      const s = v.toString().trim().toLowerCase();
      const n = parseInt(s);
      if (!isNaN(n) && n >= 1 && n <= 12) return n;
      const i1 = MONTH_NAMES.findIndex(m => m.toLowerCase() === s);
      if (i1 > 0) return i1;
      const i2 = MONTH_SHORT.findIndex(m => m.toLowerCase() === s);
      if (i2 > 0) return i2;
      return 1;
    }
    /* ---- Bantuan impor Excel/CSV: header ditulis fleksibel oleh pengguna
       (mis. "period year", "Periode (Tahun)", "Kategori"), jadi dicocokkan
       tanpa peduli spasi/huruf besar-kecil supaya format file Excel yang
       dikirim pengguna (name, category, desa, period year, month, target,
       achieved, unit) langsung terbaca tanpa perlu diubah manual. */
    function normalizedKeyMap(obj) {
      const map = {};
      Object.keys(obj || {}).forEach(k => {
        const nk = k.toString().trim().toLowerCase().replace(/[\s_()./-]+/g, '');
        map[nk] = obj[k];
      });
      return map;
    }
    function pickField(map, ...candidates) {
      for (const c of candidates) {
        const nk = c.replace(/[\s_()./-]+/g, '').toLowerCase();
        if (map[nk] !== undefined && map[nk] !== null && String(map[nk]).trim() !== '') return map[nk];
      }
      return undefined;
    }
    /* ---- Mapping kategori agar tidak ada duplikat kategori di dropdown
       (mis. "Kesling" vs "kesling " dianggap sama). Kategori yang sudah ada
       dipakai sebagai ejaan baku; kategori baru dalam satu proses impor juga
       disatukan lewat seenMap. */
    function canonicalizeCategory(raw, seenMap) {
      const clean = String(raw == null ? '' : raw).trim().replace(/\s+/g, ' ');
      if (!clean) return 'Umum';
      const key = clean.toLowerCase();
      const existing = [...new Set((state.programs || []).map(p => p.category).filter(Boolean))];
      const foundExisting = existing.find(c => c.toLowerCase() === key);
      if (foundExisting) return foundExisting;
      if (seenMap) {
        if (seenMap.has(key)) return seenMap.get(key);
        seenMap.set(key, clean);
      }
      return clean;
    }
    function uid() { return 'id-' + Math.random().toString(36).slice(2, 10); }
    function buildMonthlyProgram(name, category, arr, target, unit, period) {
      target = target || 100; unit = unit || '%'; period = period || '2026';
      return arr.map((v, i) => ({ id: uid(), name, category, period, month: i + 1, target, achieved: v, unit, desa: DESA_ALL }));
    }

    /* ---------------- DEMO DATA STORE (dipakai hanya bila DEMO_MODE) ---------------- */
    let demoDB = {
      settings: {
        theme: {
          primary: "#EC9CB6", primaryDark: "#C26A88", accent: "#F4B8CE", bg: "#FFF6F9", font: "jakarta", layout: "comfortable",
          tickerSpeed: 32,
          chart: { mode: "auto", thresholdExcellent: 100, thresholdGood: 75, colorExcellent: "#4CAF7D", colorGood: "#E8A33D", colorPoor: "#D64545", colorFixed: "#EC9CB6" }
        },
        profile: { name: "Puskesmas Puger", subtitle: "UPT Puskesmas Puger — Dinas Kesehatan Kabupaten Jember", address: "Jl. Raya Puger No. 1, Kecamatan Puger, Kabupaten Jember, Jawa Timur", phone: "(0336) 000000", emergency: "UGD 24 Jam", tagline: "Melayani Sepenuh Hati untuk Masyarakat Puger", mapsUrl: "" },
        totp_secret: ""
      },
      service_hours: DAY_NAMES.map((d, i) => ({ id: i, day_order: i, day_name: d, is_closed: i === 6, open_time: i === 6 ? null : (i === 4 ? '07:30' : '07:30'), close_time: i === 6 ? null : (i === 4 ? '11:00' : (i === 5 ? '12:30' : '14:00')), note: i === 6 ? 'Libur — Layanan UGD tetap 24 jam' : null })),
      announcements: [
        { id: 'a1', title: 'Jadwal Posyandu Balita', content: 'Posyandu balita dilaksanakan setiap tanggal 10 di masing-masing dusun.', priority: 2, is_active: true, created_at: new Date().toISOString() },
        { id: 'a2', title: 'Layanan Vaksinasi Booster', content: 'Layanan vaksinasi booster tersedia setiap Senin–Jumat, tanpa biaya.', priority: 1, is_active: true, created_at: new Date().toISOString() },
        { id: 'a3', title: 'Pemeriksaan IVA Gratis', content: 'Pemeriksaan IVA gratis bagi PUS setiap hari Rabu.', priority: 0, is_active: true, created_at: new Date().toISOString() }
      ],
      programs: [
        ...buildMonthlyProgram('Imunisasi Dasar Lengkap Balita', 'Imunisasi', [6, 13, 20, 26, 33, 40, 46, 53, 60, 66, 72, 78]),
        ...buildMonthlyProgram('Kunjungan Ibu Hamil (K4)', 'KIA', [7, 14, 22, 29, 37, 44, 52, 60, 67, 74, 80, 86]),
        ...buildMonthlyProgram('Balita Ditimbang (D/S)', 'Gizi', [8, 16, 24, 33, 41, 50, 58, 66, 74, 82, 87, 91]),
        ...buildMonthlyProgram('Penemuan Kasus TB Paru', 'P2P', [5, 10, 16, 21, 27, 33, 39, 45, 50, 56, 60, 64]),
        ...buildMonthlyProgram('Rumah Tangga ber-PHBS', 'Promkes', [6, 12, 18, 24, 30, 37, 43, 49, 55, 61, 67, 72]),
        ...buildMonthlyProgram('Akses Sanitasi Layak', 'Kesling', [7, 15, 22, 29, 37, 44, 52, 60, 67, 74, 81, 88])
      ],
      services: [
        { id: uid(), name: 'Poli Umum', category: 'Rawat Jalan', description: 'Pemeriksaan dan pengobatan penyakit umum untuk semua usia.', is_active: true },
        { id: uid(), name: 'Poli Gigi', category: 'Rawat Jalan', description: 'Pemeriksaan, penambalan, dan pencabutan gigi.', is_active: true },
        { id: uid(), name: 'KIA & KB', category: 'Kesehatan Ibu & Anak', description: 'Pemeriksaan kehamilan, imunisasi, dan pelayanan keluarga berencana.', is_active: true },
        { id: uid(), name: 'Laboratorium Sederhana', category: 'Penunjang', description: 'Pemeriksaan darah, urine, dan gula darah dasar.', is_active: true },
        { id: uid(), name: 'UGD 24 Jam', category: 'Gawat Darurat', description: 'Penanganan kegawatdaruratan medis setiap saat.', is_active: true },
        { id: uid(), name: 'Farmasi', category: 'Penunjang', description: 'Penyediaan dan penyerahan obat sesuai resep dokter.', is_active: true }
      ],
      facilities: [
        { id: uid(), name: 'Ruang Rawat Inap', category: 'Fasilitas Gedung', description: 'Tersedia tempat tidur untuk perawatan pasien rawat inap.', is_active: true },
        { id: uid(), name: 'Ambulans', category: 'Transportasi', description: 'Unit ambulans siaga untuk rujukan dan kegawatdaruratan.', is_active: true },
        { id: uid(), name: 'Ruang Bersalin', category: 'Fasilitas Gedung', description: 'Fasilitas persalinan normal dengan tenaga bidan terlatih.', is_active: true },
        { id: uid(), name: 'Apotek', category: 'Penunjang', description: 'Layanan penyediaan obat-obatan bagi pasien.', is_active: true },
        { id: uid(), name: 'Ruang Tunggu Ramah Anak', category: 'Fasilitas Gedung', description: 'Area tunggu nyaman dilengkapi permainan edukatif anak.', is_active: true },
        { id: uid(), name: 'Akses Difabel', category: 'Fasilitas Gedung', description: 'Jalur landai dan fasilitas ramah bagi penyandang disabilitas.', is_active: true }
      ],
      clusters: [
        { id: 1, title: 'Klaster 1 – Manajemen', description: 'Mengelola pendaftaran, rekam medis, kegawatdaruratan, rujukan, dan tata kelola administrasi puskesmas.', items: ['Pendaftaran & Rekam Medis', 'Pelayanan Kegawatdaruratan', 'Rujukan & Ambulans', 'Kefarmasian'] },
        { id: 2, title: 'Klaster 2 – Kesehatan Ibu, Anak & Remaja', description: 'Melayani ibu hamil, ibu bersalin, bayi, balita, anak sekolah, dan remaja termasuk imunisasi dan pemantauan tumbuh kembang.', items: ['Pemeriksaan Kehamilan (ANC)', 'Imunisasi Bayi & Balita', 'Tumbuh Kembang Anak', 'Kesehatan Reproduksi Remaja'] },
        { id: 3, title: 'Klaster 3 – Usia Dewasa & Lansia', description: 'Melayani skrining penyakit tidak menular, kesehatan jiwa, kesehatan kerja, dan pelayanan bagi lanjut usia.', items: ['Skrining Penyakit Tidak Menular', 'Kesehatan Jiwa', 'Kesehatan Kerja & Olahraga', 'Posyandu Lansia'] },
        { id: 4, title: 'Klaster 4 – Penanggulangan Penularan Penyakit', description: 'Mengendalikan penyakit menular, surveilans, dan kesehatan lingkungan di wilayah kerja puskesmas.', items: ['Pengendalian TBC & HIV', 'Surveilans Penyakit Menular', 'Kesehatan Lingkungan', 'Imunisasi & Vaksinasi Massal'] }
      ],
      staffNeeds: [
        { id: uid(), posisi: 'Bidan Desa', kualifikasi: 'D3/D4 Kebidanan, STR aktif', jumlah: 2, batas_lamar: '2026-09-15', keterangan: 'Penempatan di wilayah kerja Puskesmas Puger', is_active: true, form_file: null, form_file_name: null },
        { id: uid(), posisi: 'Tenaga Gizi', kualifikasi: 'D3 Gizi, pengalaman posyandu diutamakan', jumlah: 1, batas_lamar: '2026-09-01', keterangan: 'Status kontrak BLUD', is_active: true, form_file: null, form_file_name: null }
      ],
      applications: [],
      doctorSchedule: [
        { id: uid(), nama_dokter: 'dr. Ahmad Fauzi, Sp.PD', spesialisasi: 'Penyakit Dalam', hari: 'Selasa & Kamis', jam: '09.00–12.00', lokasi: 'Poli Spesialis', keterangan: 'Perlu pendaftaran H-1', is_active: true },
        { id: uid(), nama_dokter: 'dr. Siti Rahma, Sp.A', spesialisasi: 'Anak', hari: 'Senin & Rabu', jam: '09.00–12.00', lokasi: 'Poli Spesialis', keterangan: '', is_active: true }
      ],
      hotline: [
        { id: uid(), nama: 'UGD 24 Jam', nomor: '0336-000000', keterangan: 'Kegawatdaruratan medis', is_active: true },
        { id: uid(), nama: 'Hotline Homecare', nomor: '0812-3456-7890', keterangan: 'Layanan kunjungan rumah', is_active: true },
        { id: uid(), nama: 'Ambulans', nomor: '0336-111222', keterangan: 'Rujukan & antar jemput pasien', is_active: true }
      ],
      breakingNews: [],
      profiles: [
        { id: 'demo-admin-1', email: 'puskesmaspuger123@gmail.com', role: 'admin', is_active: true, created_at: new Date().toISOString() }
      ],
      auditLog: []
    };
    let demoAuthed = false;
    /* Menyimpan sementara data URL logo yang baru dipilih admin di form Pengaturan Tampilan,
       sebelum tombol "Simpan Profil" ditekan. undefined = tidak diubah, null = kembali ke default. */
    let pendingLogoDataUrl;

    /* ---------------- STATE ---------------- */
    let state = {
      programs: [], announcements: [], hours: [], services: [], facilities: [], clusters: [],
      staffNeeds: [], doctorSchedule: [], hotline: [], breakingNews: [], applications: [], profiles: [], settings: { theme: {}, profile: {}, totp_secret: '' }, activeCategory: '', session: null
    };

    /* ---------------- HELPERS ---------------- */
    function toast(msg, isErr) {
      const t = document.getElementById('toast');
      t.textContent = msg; t.className = 'show' + (isErr ? ' err' : '');
      clearTimeout(t._timer);
      t._timer = setTimeout(() => t.classList.remove('show'), 2600);
    }
    function pct(a, t) { if (!t) return 0; return Math.max(0, Math.min(100, Math.round((a / t) * 100))); }
    /* Membaca file gambar yang dipilih admin, lalu mengecilkannya (maks maxDim px)
       lewat <canvas> sebelum diubah jadi data URL — supaya logo hasil unggah tidak
       membengkakkan ukuran data settings meski file asli beresolusi besar. */
    function fileToResizedDataUrl(file, maxDim) {
      return new Promise((resolve, reject) => {
        if (!file.type.startsWith('image/')) { reject(new Error('File harus berupa gambar')); return; }
        if (file.size > 8 * 1024 * 1024) { reject(new Error('Ukuran file maksimal 8MB')); return; }
        const reader = new FileReader();
        reader.onerror = () => reject(new Error('Gagal membaca file'));
        reader.onload = () => {
          const img = new Image();
          img.onerror = () => reject(new Error('File bukan gambar yang valid'));
          img.onload = () => {
            let { width, height } = img;
            if (width > maxDim || height > maxDim) {
              if (width > height) { height = Math.round(height * maxDim / width); width = maxDim; }
              else { width = Math.round(width * maxDim / height); height = maxDim; }
            }
            const canvas = document.createElement('canvas');
            canvas.width = width; canvas.height = height;
            canvas.getContext('2d').drawImage(img, 0, 0, width, height);
            resolve(canvas.toDataURL('image/png'));
          };
          img.src = reader.result;
        };
        reader.readAsDataURL(file);
      });
    }
    /* Membaca berkas apa pun (PDF/DOC/gambar) menjadi data URL tanpa diubah ukurannya —
       dipakai untuk formulir pendaftaran & berkas lamaran, berbeda dari fileToResizedDataUrl
       yang khusus memampatkan gambar logo/poster. */
    function fileToDataUrl(file, maxSizeMB, allowedExt) {
      return new Promise((resolve, reject) => {
        if (allowedExt && !allowedExt.some(ext => file.name.toLowerCase().endsWith(ext))) {
          reject(new Error('Format file tidak didukung (' + allowedExt.join(', ') + ')')); return;
        }
        if (file.size > maxSizeMB * 1024 * 1024) { reject(new Error('Ukuran file maksimal ' + maxSizeMB + 'MB')); return; }
        const reader = new FileReader();
        reader.onerror = () => reject(new Error('Gagal membaca file'));
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(file);
      });
    }
    function isSafeFileSrc(u) { return !!u && u.startsWith('data:'); }

    /* ---------------- TOTP (Google Authenticator, RFC 6238) ----------------
       Diimplementasikan langsung dengan Web Crypto (HMAC-SHA1) tanpa pustaka luar,
       supaya kode verifikasi 6 digit yang tampil di aplikasi Google Authenticator
       admin bisa dicocokkan di sisi klien saat pendaftar mengunggah berkas. */
    const TOTP_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
    function base32Decode(base32) {
      let bits = '';
      const clean = (base32 || '').toUpperCase().replace(/[^A-Z2-7]/g, '');
      for (const ch of clean) {
        const val = TOTP_ALPHABET.indexOf(ch);
        if (val === -1) continue;
        bits += val.toString(2).padStart(5, '0');
      }
      const bytes = [];
      for (let i = 0; i + 8 <= bits.length; i += 8) bytes.push(parseInt(bits.substring(i, i + 8), 2));
      return new Uint8Array(bytes);
    }
    function randomBase32Secret(length) {
      length = length || 16;
      const arr = new Uint8Array(length);
      crypto.getRandomValues(arr);
      let out = '';
      for (let i = 0; i < length; i++) out += TOTP_ALPHABET[arr[i] % 32];
      return out;
    }
    async function totpGenerate(secretBase32, forTime) {
      const keyBytes = base32Decode(secretBase32);
      if (!keyBytes.length) return null;
      const counter = Math.floor((forTime ?? Date.now()) / 1000 / 30);
      const counterBuf = new ArrayBuffer(8);
      const view = new DataView(counterBuf);
      view.setUint32(0, 0, false);
      view.setUint32(4, counter, false);
      const cryptoKey = await crypto.subtle.importKey('raw', keyBytes, { name: 'HMAC', hash: 'SHA-1' }, false, ['sign']);
      const hmac = new Uint8Array(await crypto.subtle.sign('HMAC', cryptoKey, counterBuf));
      const offset = hmac[hmac.length - 1] & 0xf;
      const bin = ((hmac[offset] & 0x7f) << 24) | ((hmac[offset + 1] & 0xff) << 16) | ((hmac[offset + 2] & 0xff) << 8) | (hmac[offset + 3] & 0xff);
      return (bin % 1000000).toString().padStart(6, '0');
    }
    async function totpVerify(secretBase32, code) {
      if (!secretBase32 || !code) return false;
      const clean = String(code).trim();
      if (!/^\d{6}$/.test(clean)) return false;
      const now = Date.now();
      for (let w = -1; w <= 1; w++) {
        const otp = await totpGenerate(secretBase32, now + w * 30000);
        if (otp === clean) return true;
      }
      return false;
    }
    function escapeHtml(s) { return (s ?? '').toString().replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
    /* BUGFIX: kolom time di Postgres/Supabase mengembalikan format "HH:MM:SS" (dengan detik),
       sehingga tanpa dipotong, papan publik akan menampilkan "07:30:00–14:00:00" alih-alih "07:30–14:00". */
    function formatTime(t) { return t ? t.toString().slice(0, 5) : '-'; }
    /* BUGFIX: validasi URL sebelum dipakai sebagai href/src agar tidak rentan terhadap
       skema berbahaya seperti "javascript:" pada tautan Maps/Video yang diisi lewat admin. */
    function isSafeUrl(u) {
      if (!u) return false;
      try { const p = new URL(u, location.href); return p.protocol === 'http:' || p.protocol === 'https:'; }
      catch (e) { return false; }
    }
    /* Logo profil boleh berupa data URL (hasil unggah admin, sudah diproses lewat canvas)
       atau tautan http/https biasa. */
    function isSafeImageSrc(u) {
      if (!u) return false;
      if (u.startsWith('data:image/')) return true;
      return isSafeUrl(u);
    }
    function applyProfileLogo(url) {
      const src = isSafeImageSrc(url) ? url : DEFAULT_PROFILE_LOGO;
      const a = document.getElementById('pf-logo-img'); if (a) a.src = src;
      const b = document.getElementById('admin-logo-img'); if (b) b.src = src;
    }

    /* ---------------- THEME APPLY ---------------- */
    function applyTheme(theme) {
      const r = document.documentElement.style;
      if (theme.primary) r.setProperty('--primary', theme.primary);
      if (theme.primaryDark) r.setProperty('--primary-dark', theme.primaryDark);
      if (theme.accent) r.setProperty('--accent', theme.accent);
      if (theme.bg) r.setProperty('--bg', theme.bg);
      if (theme.tickerSpeed) r.setProperty('--ticker-duration', theme.tickerSpeed + 's');
      document.body.classList.toggle('layout-compact', theme.layout === 'compact');
    }

    const THEME_PRESETS = [
      { name: 'Pink Lembut', primary: '#EC9CB6', primaryDark: '#C26A88', accent: '#F4B8CE', bg: '#FFF6F9' },
      { name: 'Pink Mawar', primary: '#E17FA0', primaryDark: '#B4506F', accent: '#F7C9DA', bg: '#FFF3F6' },
      { name: 'Pink Lavender', primary: '#D98FC0', primaryDark: '#A85B93', accent: '#F0C9E6', bg: '#FCF5FB' },
      { name: 'Pink Peach', primary: '#EFA6A0', primaryDark: '#C56E68', accent: '#FBD3C6', bg: '#FFF7F4' },
    ];

    /* ---------------- WARNA GRAFIK BERDASARKAN CAPAIAN ---------------- */
    function defaultChartSettings() {
      return { mode: 'auto', thresholdExcellent: 100, thresholdGood: 75, colorExcellent: '#4CAF7D', colorGood: '#E8A33D', colorPoor: '#D64545', colorFixed: '#EC9CB6' };
    }
    function achievementColor(pctValue, themeChart, fallbackPrimary) {
      const c = { ...defaultChartSettings(), ...(themeChart || {}) };
      if (c.mode === 'fixed') return c.colorFixed || fallbackPrimary;
      if (pctValue >= c.thresholdExcellent) return c.colorExcellent;
      if (pctValue >= c.thresholdGood) return c.colorGood;
      return c.colorPoor;
    }

    /* ---------------- LINE CHART (SVG) ---------------- */
    function groupProgramSeries(rows) {
      const map = new Map();
      /* Grafik & angka total puskesmas hanya dihitung dari baris agregat ("Semua Desa"),
         supaya baris rincian per desa (dipilih lewat menu Desa di form edit) tidak
         ikut menumpuk/menimpa angka total bulan yang sama. */
      rows.filter(r => isAggregateDesa(r.desa)).forEach(r => {
        const key = r.name + '|' + r.category + '|' + r.period;
        if (!map.has(key)) map.set(key, { name: r.name, category: r.category, period: r.period, unit: r.unit, target: Number(r.target) || 0, monthly: {} });
        const g = map.get(key);
        /* BUGFIX: Number(...) wajib di sini — kolom "numeric" Postgres dikirim PostgREST sebagai
           string (mis. "78.00"), sehingga `sum + r.achieved` sebelumnya melakukan penggabungan
           teks alih-alih penjumlahan dan membuat avg.toFixed() error di produksi (non-demo). */
        g.monthly[r.month] = Number(r.achieved) || 0;
        if (r.target) g.target = Number(r.target) || 0;
        if (r.unit) g.unit = r.unit;
      });
      return [...map.values()];
    }

    /* Hash string sederhana -> angka 0..1, dipakai agar variasi capaian
       tiap desa konsisten setiap kali dibuka (bukan acak ulang tiap render). */
    function seededRatio(str) {
      let h = 0;
      for (let i = 0; i < str.length; i++) { h = (h << 5) - h + str.charCodeAt(i); h |= 0; }
      return ((h % 1000) + 1000) % 1000 / 1000;
    }

    /* Rincian capaian per desa untuk satu series program — dihitung OTOMATIS,
       per desa, tanpa mengharuskan admin mengisi capaian tiap desa satu per satu:
       - Desa yang SUDAH punya data ASLI (diinput admin lewat menu Desa di form
         edit "Capaian Layanan", baris dengan desa = salah satu dari DESA_LIST)
         memakai data asli tersebut (isEstimated: false).
       - Desa yang BELUM diisi otomatis diperkirakan (direkap) dari angka agregat
         Puskesmas ("Semua Desa") yang sudah diinput, supaya papan publik tetap
         menampilkan rincian per desa yang wajar walau admin baru sempat mengisi
         sebagian desa saja (isEstimated: true).
       Kedua mode digabung per desa (bukan semua-atau-tidak-sama-sekali), jadi
       mengisi satu desa tidak membuat desa lain jadi kosong/0%. */
    function villageBreakdownFor(s) {
      const months = Object.keys(s.monthly).map(Number).sort((a, b) => a - b);
      const realRows = (state.programs || []).filter(r =>
        r.name === s.name && r.category === s.category && r.period === s.period && DESA_LIST.includes(r.desa));
      return DESA_LIST.map(desa => {
        const rowsForDesa = realRows.filter(r => r.desa === desa);
        if (rowsForDesa.length) {
          const monthly = {};
          rowsForDesa.forEach(r => { monthly[r.month] = Number(r.achieved) || 0; });
          const desaMonths = Object.keys(monthly).map(Number).sort((a, b) => a - b);
          const latestMonth = desaMonths.length ? desaMonths[desaMonths.length - 1] : null;
          const latestVal = latestMonth ? monthly[latestMonth] : 0;
          const avg = desaMonths.length ? desaMonths.reduce((sum, m) => sum + monthly[m], 0) / desaMonths.length : 0;
          return { desa, monthly, latestMonth, latestVal, avg, isEstimated: false };
        }
        const r = seededRatio(s.name + '|' + s.category + '|' + desa);
        const factor = 0.82 + r * 0.34; // sebaran ±18% dari capaian puskesmas (perkiraan otomatis)
        const monthly = {};
        months.forEach(m => { monthly[m] = Math.max(0, Math.round(s.monthly[m] * factor * 10) / 10); });
        const latestMonth = months.length ? months[months.length - 1] : null;
        const latestVal = latestMonth ? monthly[latestMonth] : 0;
        const avg = months.length ? months.reduce((sum, m) => sum + monthly[m], 0) / months.length : 0;
        return { desa, monthly, latestMonth, latestVal, avg, isEstimated: true };
      });
    }
    function lineChartSVG(monthly, target, unit, lineColor) {
      const W = 440, H = 160, padL = 26, padR = 10, padT = 14, padB = 24;
      const color = lineColor || 'var(--primary)';
      const months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
      const values = months.map(m => monthly[m]).filter(v => v != null);
      const maxVal = Math.max(target || 0, ...values, 10) * 1.15;
      const x = m => padL + (m - 1) * ((W - padL - padR) / 11);
      const y = v => H - padB - (v / maxVal) * (H - padT - padB);
      const pts = months.filter(m => monthly[m] != null).map(m => `${x(m)},${y(monthly[m])}`).join(' ');
      const dots = months.filter(m => monthly[m] != null).map(m =>
        `<circle cx="${x(m)}" cy="${y(monthly[m])}" r="3.2" fill="${color}"><title>${MONTH_SHORT[m]}: ${monthly[m]}${unit || ''}</title></circle>`
      ).join('');
      const targetY = y(target || 0);
      const monthLabels = months.map(m => `<text x="${x(m)}" y="${H - 8}" font-size="8.5" fill="var(--muted)" text-anchor="middle">${MONTH_SHORT[m]}</text>`).join('');
      return `<svg viewBox="0 0 ${W} ${H}" class="line-chart-svg" preserveAspectRatio="xMidYMid meet">
    <line x1="${padL}" y1="${targetY}" x2="${W - padR}" y2="${targetY}" stroke="var(--accent-dark)" stroke-width="1.4" stroke-dasharray="4 4"/>
    <polyline points="${pts}" fill="none" stroke="${color}" stroke-width="2.4" stroke-linejoin="round" stroke-linecap="round"/>
    ${dots}
    ${monthLabels}
  </svg>`;
    }

    /* =====================================================================
       DATA LAYER — abstraksi Supabase / demo agar logika UI seragam
       ===================================================================== */
    const DataLayer = {
      async getSettings() {
        if (DEMO_MODE) return demoDB.settings;
        const { data, error } = await sb.from('settings').select('*');
        if (error) throw error;
        const out = { theme: {}, profile: {} };
        (data || []).forEach(r => out[r.key] = r.value);
        return out;
      },
      async saveSetting(key, value) {
        if (DEMO_MODE) { demoDB.settings[key] = value; return; }
        const { error } = await sb.from('settings').upsert({ key, value, updated_at: new Date().toISOString() });
        if (error) throw error;
      },
      async getHours() {
        if (DEMO_MODE) return [...demoDB.service_hours].sort((a, b) => a.day_order - b.day_order);
        const { data, error } = await sb.from('service_hours').select('*').order('day_order');
        if (error) throw error;
        return data;
      },
      async saveHour(row) {
        if (DEMO_MODE) {
          const i = demoDB.service_hours.findIndex(h => h.day_order === row.day_order);
          if (i > -1) demoDB.service_hours[i] = { ...demoDB.service_hours[i], ...row };
          return;
        }
        const { error } = await sb.from('service_hours').update(row).eq('day_order', row.day_order);
        if (error) throw error;
      },
      async getAnnouncements(all) {
        if (DEMO_MODE) return all ? demoDB.announcements : demoDB.announcements.filter(a => a.is_active);
        let q = sb.from('announcements').select('*').order('priority', { ascending: false });
        if (!all) q = q.eq('is_active', true);
        const { data, error } = await q;
        if (error) throw error;
        return data;
      },
      async upsertAnnouncement(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); row.created_at = new Date().toISOString(); demoDB.announcements.unshift(row); }
          else { const i = demoDB.announcements.findIndex(a => a.id === row.id); demoDB.announcements[i] = row; }
          return;
        }
        const { error } = await sb.from('announcements').upsert(row);
        if (error) throw error;
      },
      async deleteAnnouncement(id) {
        if (DEMO_MODE) { demoDB.announcements = demoDB.announcements.filter(a => a.id !== id); return; }
        const { error } = await sb.from('announcements').delete().eq('id', id);
        if (error) throw error;
      },
      async getPrograms() {
        if (DEMO_MODE) return [...demoDB.programs];
        const { data, error } = await sb.from('programs').select('*').order('period', { ascending: false }).order('month');
        if (error) throw error;
        return data;
      },
      async upsertProgram(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); demoDB.programs.unshift(row); }
          else { const i = demoDB.programs.findIndex(p => p.id === row.id); demoDB.programs[i] = row; }
          return;
        }
        const { error } = await sb.from('programs').upsert(row);
        if (error) throw error;
      },
      async deleteProgram(id) {
        if (DEMO_MODE) { demoDB.programs = demoDB.programs.filter(p => p.id !== id); return; }
        const { error } = await sb.from('programs').delete().eq('id', id);
        if (error) throw error;
      },
      async bulkImportPrograms(rows) {
        if (DEMO_MODE) {
          rows.forEach(r => {
            const rDesa = r.desa || DESA_ALL;
            const existing = demoDB.programs.find(p => p.name === r.name && p.category === r.category && p.period === r.period && p.month === r.month && (p.desa || DESA_ALL) === rDesa);
            if (existing) Object.assign(existing, r, { desa: rDesa });
            else demoDB.programs.unshift({ id: uid(), ...r, desa: rDesa });
          });
          return;
        }
        for (const r of rows) {
          const rDesa = r.desa || DESA_ALL;
          const { data: existing } = await sb.from('programs').select('id').eq('name', r.name).eq('category', r.category).eq('period', r.period).eq('month', r.month).eq('desa', rDesa).maybeSingle();
          const payload = { ...r, desa: rDesa };
          if (existing) payload.id = existing.id;
          const { error } = await sb.from('programs').upsert(payload);
          if (error) throw error;
        }
      },

      async getServices() {
        if (DEMO_MODE) return [...demoDB.services];
        const { data, error } = await sb.from('services').select('*');
        if (error) throw error;
        return data;
      },
      async upsertService(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); demoDB.services.unshift(row); }
          else { const i = demoDB.services.findIndex(x => x.id === row.id); demoDB.services[i] = row; }
          return;
        }
        const { error } = await sb.from('services').upsert(row);
        if (error) throw error;
      },
      async deleteService(id) {
        if (DEMO_MODE) { demoDB.services = demoDB.services.filter(x => x.id !== id); return; }
        const { error } = await sb.from('services').delete().eq('id', id);
        if (error) throw error;
      },

      async getFacilities() {
        if (DEMO_MODE) return [...demoDB.facilities];
        const { data, error } = await sb.from('facilities').select('*');
        if (error) throw error;
        return data;
      },
      async upsertFacility(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); demoDB.facilities.unshift(row); }
          else { const i = demoDB.facilities.findIndex(x => x.id === row.id); demoDB.facilities[i] = row; }
          return;
        }
        const { error } = await sb.from('facilities').upsert(row);
        if (error) throw error;
      },
      async deleteFacility(id) {
        if (DEMO_MODE) { demoDB.facilities = demoDB.facilities.filter(x => x.id !== id); return; }
        const { error } = await sb.from('facilities').delete().eq('id', id);
        if (error) throw error;
      },

      async getClusters() {
        if (DEMO_MODE) return [...demoDB.clusters].sort((a, b) => a.id - b.id);
        const { data, error } = await sb.from('clusters').select('*').order('id');
        if (error) throw error;
        return data;
      },
      /* updateCluster: dipertahankan untuk menyimpan perubahan klaster yang sudah ada (dipanggil
         dengan row.id terisi). Menambah/menghapus klaster memakai upsertCluster/deleteCluster. */
      async updateCluster(row) {
        if (DEMO_MODE) {
          const i = demoDB.clusters.findIndex(c => c.id === row.id);
          if (i > -1) demoDB.clusters[i] = { ...demoDB.clusters[i], ...row };
          return;
        }
        const { error } = await sb.from('clusters').update(row).eq('id', row.id);
        if (error) throw error;
      },
      async upsertCluster(row) {
        if (DEMO_MODE) {
          if (row.id === undefined || row.id === null) {
            const nextId = demoDB.clusters.reduce((m, c) => Math.max(m, c.id || 0), 0) + 1;
            row = { ...row, id: nextId };
            demoDB.clusters.push(row);
          } else {
            const i = demoDB.clusters.findIndex(c => c.id === row.id);
            if (i > -1) demoDB.clusters[i] = { ...demoDB.clusters[i], ...row };
            else demoDB.clusters.push(row);
          }
          return;
        }
        const { error } = await sb.from('clusters').upsert(row);
        if (error) throw error;
      },
      async deleteCluster(id) {
        if (DEMO_MODE) { demoDB.clusters = demoDB.clusters.filter(x => x.id !== id); return; }
        const { error } = await sb.from('clusters').delete().eq('id', id);
        if (error) throw error;
      },

      async getStaffNeeds() {
        if (DEMO_MODE) return [...demoDB.staffNeeds];
        const { data, error } = await sb.from('staff_needs').select('*');
        if (error) throw error;
        return data;
      },
      async upsertStaffNeed(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); demoDB.staffNeeds.unshift(row); }
          else { const i = demoDB.staffNeeds.findIndex(x => x.id === row.id); demoDB.staffNeeds[i] = row; }
          return;
        }
        const { error } = await sb.from('staff_needs').upsert(row);
        if (error) throw error;
      },
      async deleteStaffNeed(id) {
        if (DEMO_MODE) { demoDB.staffNeeds = demoDB.staffNeeds.filter(x => x.id !== id); return; }
        const { error } = await sb.from('staff_needs').delete().eq('id', id);
        if (error) throw error;
      },

      async getDoctorSchedule() {
        if (DEMO_MODE) return [...demoDB.doctorSchedule];
        const { data, error } = await sb.from('doctor_schedule').select('*');
        if (error) throw error;
        return data;
      },
      async upsertDoctorSchedule(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); demoDB.doctorSchedule.unshift(row); }
          else { const i = demoDB.doctorSchedule.findIndex(x => x.id === row.id); demoDB.doctorSchedule[i] = row; }
          return;
        }
        const { error } = await sb.from('doctor_schedule').upsert(row);
        if (error) throw error;
      },
      async deleteDoctorSchedule(id) {
        if (DEMO_MODE) { demoDB.doctorSchedule = demoDB.doctorSchedule.filter(x => x.id !== id); return; }
        const { error } = await sb.from('doctor_schedule').delete().eq('id', id);
        if (error) throw error;
      },

      async getHotline() {
        if (DEMO_MODE) return [...demoDB.hotline];
        const { data, error } = await sb.from('hotline').select('*');
        if (error) throw error;
        return data;
      },
      async upsertHotline(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); demoDB.hotline.unshift(row); }
          else { const i = demoDB.hotline.findIndex(x => x.id === row.id); demoDB.hotline[i] = row; }
          return;
        }
        const { error } = await sb.from('hotline').upsert(row);
        if (error) throw error;
      },
      async deleteHotline(id) {
        if (DEMO_MODE) { demoDB.hotline = demoDB.hotline.filter(x => x.id !== id); return; }
        const { error } = await sb.from('hotline').delete().eq('id', id);
        if (error) throw error;
      },

      async getBreakingNews() {
        if (DEMO_MODE) return [...demoDB.breakingNews];
        const { data, error } = await sb.from('breaking_news').select('*');
        if (error) throw error;
        return data;
      },
      async upsertBreakingNews(row) {
        if (DEMO_MODE) {
          if (!row.id) { row.id = uid(); demoDB.breakingNews.unshift(row); }
          else { const i = demoDB.breakingNews.findIndex(x => x.id === row.id); demoDB.breakingNews[i] = row; }
          return;
        }
        const { error } = await sb.from('breaking_news').upsert(row);
        if (error) throw error;
      },
      async deleteBreakingNews(id) {
        if (DEMO_MODE) { demoDB.breakingNews = demoDB.breakingNews.filter(x => x.id !== id); return; }
        const { error } = await sb.from('breaking_news').delete().eq('id', id);
        if (error) throw error;
      },

      async getApplications() {
        if (DEMO_MODE) return [...demoDB.applications];
        const { data, error } = await sb.from('applications').select('*').order('created_at', { ascending: false });
        if (error) throw error;
        return data;
      },
      async insertApplication(row) {
        if (DEMO_MODE) { row.id = uid(); demoDB.applications.unshift(row); return; }
        const { error } = await sb.from('applications').insert(row);
        if (error) throw error;
      },
      async deleteApplication(id) {
        if (DEMO_MODE) { demoDB.applications = demoDB.applications.filter(x => x.id !== id); return; }
        const { error } = await sb.from('applications').delete().eq('id', id);
        if (error) throw error;
      },
      /* ---- Manajemen Akun: Role & Nonaktifkan Akun ---- */
      async getProfiles() {
        if (DEMO_MODE) return [...demoDB.profiles];
        const { data, error } = await sb.from('admin_profiles').select('*').order('created_at');
        if (error) throw error;
        return data;
      },
      async getProfileById(id) {
        if (DEMO_MODE) return demoDB.profiles.find(p => p.id === id) || null;
        const { data, error } = await sb.from('admin_profiles').select('*').eq('id', id).maybeSingle();
        if (error) throw error;
        return data;
      },
      async ensureOwnProfile(id, email) {
        /* Self-heal: akun yang dibuat sebelum fitur ini ada belum punya baris di
           admin_profiles. Saat login, buat baris untuk dirinya sendiri jika belum ada. */
        if (DEMO_MODE) {
          if (!demoDB.profiles.some(p => p.id === id)) {
            demoDB.profiles.push({ id, email, role: 'admin', is_active: true, created_at: new Date().toISOString() });
          }
          return;
        }
        const { error } = await sb.from('admin_profiles')
          .upsert({ id, email, role: 'admin', is_active: true }, { onConflict: 'id', ignoreDuplicates: true });
        if (error) throw error;
      },
      async createProfile(id, email, role) {
        if (DEMO_MODE) { demoDB.profiles.push({ id, email, role: role || 'admin', is_active: true, created_at: new Date().toISOString() }); return; }
        const { error } = await sb.from('admin_profiles').insert({ id, email, role: role || 'admin', is_active: true });
        if (error) throw error;
      },
      async updateProfileRole(id, role) {
        if (DEMO_MODE) { const p = demoDB.profiles.find(x => x.id === id); if (p) p.role = role; return; }
        const { error } = await sb.from('admin_profiles').update({ role }).eq('id', id);
        if (error) throw error;
      },
      async updateProfileActive(id, is_active) {
        if (DEMO_MODE) { const p = demoDB.profiles.find(x => x.id === id); if (p) p.is_active = is_active; return; }
        const { error } = await sb.from('admin_profiles').update({ is_active }).eq('id', id);
        if (error) throw error;
      },
      /* ---- Riwayat Aktivitas Akun (audit log) ----
         Kegagalan menyimpan/mengambil log TIDAK BOLEH menggagalkan aksi utamanya
         (ganti peran/nonaktifkan/tambah akun/reset password) — log hanya bersifat
         pencatatan, bukan hal yang divalidasi. Karena itu logAudit() sengaja tidak
         melempar error ke pemanggilnya. */
      async logAudit(action, targetEmail, detail) {
        const actorEmail = (state.session && state.session.user && state.session.user.email) || '';
        const actorId = (state.session && state.session.user && state.session.user.id) || null;
        if (DEMO_MODE) {
          demoDB.auditLog.unshift({ id: 'log-' + Date.now(), actor_id: actorId, actor_email: actorEmail, action, target_email: targetEmail || '', detail: detail || '', created_at: new Date().toISOString() });
          return;
        }
        try {
          await sb.from('admin_audit_log').insert({ actor_id: actorId, actor_email: actorEmail, action, target_email: targetEmail || '', detail: detail || '' });
        } catch (e) { console.warn('Gagal mencatat riwayat aktivitas:', errMsg(e)); }
      },
      async getAuditLog(limit) {
        if (DEMO_MODE) return demoDB.auditLog.slice(0, limit || 30);
        const { data, error } = await sb.from('admin_audit_log').select('*').order('created_at', { ascending: false }).limit(limit || 30);
        if (error) throw error;
        return data;
      }
    };

    /* =====================================================================
       ROUTER
       ===================================================================== */
    function currentRoute() { return location.hash.replace('#', '') || '/'; }

    async function router() {
      const route = currentRoute();
      document.getElementById('view-public').classList.add('hidden');
      document.getElementById('view-login').classList.add('hidden');
      document.getElementById('view-admin').classList.add('hidden');

      const authed = DEMO_MODE ? demoAuthed : !!state.session;

      if (route.startsWith('/admin')) {
        if (!authed) { location.hash = '#/login'; return; }
        /* BUGFIX: state.myProfile (dipakai untuk menentukan peran Admin/Operator) sebelumnya
           hanya diisi di dalam handler submit form login. Kalau halaman dibuka langsung ke
           #/admin dengan sesi yang sudah ada (mis. reload halaman), state.myProfile kosong dan
           akun Operator keliru dianggap Admin (fallback bawaan). Muat/segarkan di sini juga
           sehingga peran & status aktif akun selalu diperiksa ulang tiap kali panel admin dibuka. */
        const keepLoggedIn = await ensureMyProfileLoaded();
        if (!keepLoggedIn) { location.hash = '#/login'; return; }
        document.getElementById('view-admin').classList.remove('hidden');
        await loadAdminData();
        const tab = route.split('/')[2] || 'ringkasan';
        setAdminTab(tab);
      } else if (route === '/login') {
        if (authed) { location.hash = '#/admin'; return; }
        document.getElementById('view-login').classList.remove('hidden');
        document.getElementById('demo-note').classList.toggle('hidden', !DEMO_MODE);
      } else {
        document.getElementById('view-public').classList.remove('hidden');
        await loadPublicData();
      }
    }
    window.addEventListener('hashchange', router);

    /* =====================================================================
       PUBLIC VIEW
       ===================================================================== */
    async function loadPublicData() {
      try {
        const [settings, hours, ann, programs, services, facilities, clusters, staffNeeds, doctorSchedule, hotline, breakingNews] = await Promise.all([
          DataLayer.getSettings(), DataLayer.getHours(), DataLayer.getAnnouncements(false), DataLayer.getPrograms(),
          DataLayer.getServices(), DataLayer.getFacilities(), DataLayer.getClusters(), DataLayer.getStaffNeeds(), DataLayer.getDoctorSchedule(), DataLayer.getHotline(), DataLayer.getBreakingNews()
        ]);
        state.settings = settings; state.hours = hours; state.announcements = ann; state.programs = programs;
        state.services = services; state.facilities = facilities; state.clusters = clusters;
        state.staffNeeds = staffNeeds; state.doctorSchedule = doctorSchedule; state.hotline = hotline; state.breakingNews = breakingNews;

        applyTheme(settings.theme || {});
        renderProfile(settings.profile || {});
        renderTicker(ann);
        renderHours(hours);
        renderCategoryFilter(programs);
        renderPrograms(programs);
        renderServices(services);
        renderFacilities(facilities);
        renderClusters(clusters);
        renderStaffNeeds(staffNeeds);
        renderDoctorSchedule(doctorSchedule);
        renderHotline(hotline);
        renderBreakingNews(breakingNews);
      } catch (e) { console.error(e); toast('Gagal memuat data: ' + errMsg(e), true); }
    }

    function renderProfile(p) {
      document.getElementById('pf-name').textContent = p.name || 'Puskesmas Puger';
      document.getElementById('pf-subtitle').textContent = p.subtitle || '';
      applyProfileLogo(p.logo);
      /* BUGFIX/ENHANCEMENT: sebelumnya textContent menimpa <span class="hl-word"> statis di HTML
         setiap kali data dimuat, sehingga efek sorot kata pada judul hilang begitu tagline diisi
         dari settings. Sekarang efek sorot dibangun ulang secara dinamis (2 kata terakhir),
         tetap aman dari XSS karena teksnya di-escape lebih dulu. */
      const tagline = p.tagline || 'Melayani Sepenuh Hati untuk Masyarakat Puger';
      document.getElementById('pf-tagline').innerHTML = highlightLastWords(tagline, 2);
      document.getElementById('pf-address').textContent = p.address || '';
      document.getElementById('pf-phone').textContent = p.phone ? ('Telepon: ' + p.phone) : '';
      document.getElementById('pf-emergency').textContent = p.emergency || '';
      const mapsEl = document.getElementById('pf-maps');
      if (p.mapsUrl && isSafeUrl(p.mapsUrl)) { mapsEl.href = p.mapsUrl; mapsEl.classList.remove('hidden'); }
      else { mapsEl.classList.add('hidden'); }
    }
    function highlightLastWords(text, n) {
      const words = (text || '').trim().split(/\s+/).filter(Boolean);
      if (!words.length) return '';
      if (words.length <= n) return `<span class="hl-word">${escapeHtml(words.join(' '))}</span>`;
      const head = words.slice(0, words.length - n).join(' ');
      const tail = words.slice(words.length - n).join(' ');
      return `${escapeHtml(head)} <span class="hl-word">${escapeHtml(tail)}</span>`;
    }

    function renderTicker(ann) {
      const el = document.getElementById('ticker-content');
      if (!ann.length) { el.innerHTML = '<div class="ticker-empty">Belum ada pengumuman.</div>'; return; }
      const items = ann.map(a => `<span>${escapeHtml(a.title)}${a.content ? (' — ' + escapeHtml(a.content)) : ''}</span>`).join('');
      el.innerHTML = `<div class="ticker-track">${items}${items}</div>`;
    }

    function renderHours(hours) {
      const jsDay = new Date().getDay(); // 0=Minggu
      const todayOrder = (jsDay + 6) % 7; // 0=Senin ... 6=Minggu
      const el = document.getElementById('hours-grid');
      el.innerHTML = hours.map(h => `
    <div class="hour-card ${h.day_order === todayOrder ? 'today' : ''}">
      <div class="day">${h.day_name}</div>
      ${h.is_closed
          ? `<div class="closed">Tutup</div>`
          : `<div class="time">${formatTime(h.open_time)}–${formatTime(h.close_time)}</div>`}
      ${h.note ? `<div class="note">${escapeHtml(h.note)}</div>` : ''}
    </div>`).join('');
    }

    function renderCategoryFilter(programs) {
      const cats = ['', ...new Set(programs.map(p => p.category))];
      const el = document.getElementById('category-filter');
      el.innerHTML = cats.map(c => `<button class="pill ${state.activeCategory === c ? 'active' : ''}" data-cat="${escapeHtml(c)}">${c || 'Semua'}</button>`).join('');
      el.querySelectorAll('.pill').forEach(b => b.onclick = () => { state.activeCategory = b.dataset.cat; renderCategoryFilter(state.programs); renderPrograms(state.programs); });
    }

    function renderPrograms(programs) {
      const list = state.activeCategory ? programs.filter(p => p.category === state.activeCategory) : programs;
      const series = groupProgramSeries(list);
      const el = document.getElementById('capaian-grid');
      const chartSettings = (state.settings.theme && state.settings.theme.chart) || defaultChartSettings();
      const legendEl = document.getElementById('capaian-legend');
      if (legendEl) {
        legendEl.innerHTML = chartSettings.mode === 'fixed' ? '' : `
      <span><i style="background:${chartSettings.colorExcellent}"></i>Baik (≥${chartSettings.thresholdExcellent}%)</span>
      <span><i style="background:${chartSettings.colorGood}"></i>Cukup (≥${chartSettings.thresholdGood}%)</span>
      <span><i style="background:${chartSettings.colorPoor}"></i>Kurang (&lt;${chartSettings.thresholdGood}%)</span>`;
      }
      if (!series.length) { el.innerHTML = '<p style="color:var(--muted);">Belum ada data capaian.</p>'; return; }
      /* disimpan agar openVillageModal() bisa mengambil series yang sedang
         ditampilkan (termasuk pengaturan warna) saat kartu diklik publik */
      state.lastCapaianSeries = series;
      state.lastChartSettings = chartSettings;
      el.innerHTML = series.map((s, idx) => {
        const months = Object.keys(s.monthly).map(Number);
        const latestMonth = months.length ? Math.max(...months) : null;
        const latestVal = latestMonth ? s.monthly[latestMonth] : null;
        const avg = months.length ? (months.reduce((sum, m) => sum + s.monthly[m], 0) / months.length) : 0;
        const latestPct = latestVal != null ? pct(latestVal, s.target) : 0;
        const lineColor = achievementColor(latestPct, chartSettings, 'var(--primary)');
        return `<div class="chart-card" style="border-left-color:${lineColor}" role="button" tabindex="0"
      aria-label="Lihat capaian ${escapeHtml(s.name)} per desa"
      onclick="openVillageModal(${idx})" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();openVillageModal(${idx});}">
      <div class="chead">
        <div>
          <span class="program-cat">${escapeHtml(s.category)}</span>
          <h4>${escapeHtml(s.name)}</h4>
        </div>
        <div class="cur">
          <div class="v" style="color:${lineColor}">${latestPct}%</div>
          <div class="m">${latestMonth ? MONTH_SHORT[latestMonth] : ''} ${escapeHtml(s.period)}</div>
        </div>
      </div>
      ${lineChartSVG(s.monthly, s.target, s.unit, lineColor)}
      <div class="stats">Rata-rata capaian <b>${avg.toFixed(1)}${s.unit || ''}</b> · Target <b>${s.target}${s.unit || ''}</b></div>
      <div class="click-hint">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="7" r="3"/><path d="M2 21c0-4 3-6 7-6s7 2 7 6"/><circle cx="17" cy="8" r="2.3"/><path d="M15.5 15c3 .3 5 2 5 6"/></svg>
        Klik untuk lihat capaian per desa
      </div>
    </div>`;
      }).join('');
    }

    /* ---------------- MODAL: RINCIAN CAPAIAN PER DESA (papan publik) ---------------- */
    function openVillageModal(idx) {
      const series = state.lastCapaianSeries || [];
      const s = series[idx];
      if (!s) return;
      const chartSettings = state.lastChartSettings;
      const months = Object.keys(s.monthly).map(Number);
      const latestMonth = months.length ? Math.max(...months) : null;
      const latestVal = latestMonth ? s.monthly[latestMonth] : null;
      const overallPct = latestVal != null ? pct(latestVal, s.target) : 0;
      const overallColor = achievementColor(overallPct, chartSettings, 'var(--primary-dark)');

      const desaRows = villageBreakdownFor(s)
        .map(v => ({ ...v, p: pct(v.latestVal, s.target) }))
        .sort((a, b) => b.p - a.p);

      const rowsHtml = desaRows.map(v => {
        const color = achievementColor(v.p, chartSettings, 'var(--primary-dark)');
        const badge = v.isEstimated
          ? `<span class="badge off" title="Belum diinput admin untuk desa ini — dihitung otomatis dari capaian total Puskesmas">≈ perkiraan otomatis</span>`
          : `<span class="badge" title="Diisi langsung oleh admin untuk desa ini">data terinput</span>`;
        return `<div class="vm-desa-row">
        <div class="vm-desa-top">
          <span class="vm-desa-name">Desa ${escapeHtml(v.desa)} ${badge}</span>
          <span class="vm-desa-pct" style="color:${color}">${v.p}%</span>
        </div>
        <div class="vm-bar-track"><div class="vm-bar-fill" style="width:${Math.min(100, v.p)}%;background:${color}"></div></div>
        <div class="stats" style="margin-top:6px;">Capaian ${v.latestVal}${s.unit || ''} dari target ${s.target}${s.unit || ''} · Rata-rata ${v.avg.toFixed(1)}${s.unit || ''}</div>
      </div>`;
      }).join('');

      const root = document.getElementById('modal-root');
      root.innerHTML = `
  <div class="modal-overlay" onclick="if(event.target===this) closeModal()">
    <div class="modal village-modal">
      <div class="vm-head">
        <div>
          <span class="program-cat">${escapeHtml(s.category)}</span>
          <h3 style="margin-top:6px;">${escapeHtml(s.name)}</h3>
          <p class="vm-sub">Rincian capaian per desa binaan wilayah kerja Puskesmas Puger — ${latestMonth ? MONTH_SHORT[latestMonth] : ''} ${escapeHtml(s.period)}</p>
        </div>
        <button class="vm-close" onclick="closeModal()" aria-label="Tutup">✕</button>
      </div>
      <div class="vm-overall">
        <span class="lbl">Capaian Puskesmas Puger (keseluruhan)</span>
        <span class="val" style="color:${overallColor}">${overallPct}%</span>
      </div>
      <div class="vm-desa-list">${rowsHtml}</div>
      <p class="vm-foot">Data dikelompokkan per desa untuk memudahkan pemantauan capaian program kesehatan oleh warga dan kader di wilayah masing-masing. Target capaian program ini adalah ${s.target}${s.unit || ''}. Desa dengan label <b>"≈ perkiraan otomatis"</b> belum diinput datanya secara khusus oleh admin, sehingga sistem menghitungnya secara otomatis (direkap) dari capaian total Puskesmas.</p>
    </div>
  </div>`;
    }
    /* memudahkan penutupan modal dengan tombol Escape, termasuk untuk modal rincian desa */
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeModal(); });

    /* Data aktif disimpan di variabel global supaya bisa dirujuk lewat index
       saat kartu diklik, tanpa perlu menyisipkan seluruh objek ke atribut onclick. */
    let __servicesActive = [];
    let __facilitiesActive = [];
    let __clustersActive = [];
    let __staffActive = [];
    let __doctorsActive = [];

    function renderServices(list) {
      const el = document.getElementById('layanan-grid');
      const active = list.filter(s => s.is_active);
      __servicesActive = active;
      el.innerHTML = active.length ? active.map((s, i) => `
    <div class="simple-card" onclick="openServiceDetail(${i})" role="button" tabindex="0" onkeypress="if(event.key==='Enter') openServiceDetail(${i})">
      <span class="program-cat">${escapeHtml(s.category)}</span>
      <h4>${escapeHtml(s.name)}</h4>
      <p>${escapeHtml(s.description)}</p>
    </div>`).join('') : '<p style="color:var(--muted);">Belum ada informasi layanan.</p>';
    }

    function openServiceDetail(i) {
      const s = __servicesActive[i];
      if (!s) return;
      openSimpleDetailModal(s.category, s.name, s.description);
    }

    function renderFacilities(list) {
      const el = document.getElementById('fasilitas-grid');
      const active = list.filter(s => s.is_active);
      __facilitiesActive = active;
      el.innerHTML = active.length ? active.map((s, i) => `
    <div class="simple-card" onclick="openFacilityDetail(${i})" role="button" tabindex="0" onkeypress="if(event.key==='Enter') openFacilityDetail(${i})">
      <span class="program-cat">${escapeHtml(s.category)}</span>
      <h4>${escapeHtml(s.name)}</h4>
      <p>${escapeHtml(s.description)}</p>
    </div>`).join('') : '<p style="color:var(--muted);">Belum ada data fasilitas.</p>';
    }

    function openFacilityDetail(i) {
      const s = __facilitiesActive[i];
      if (!s) return;
      openSimpleDetailModal(s.category, s.name, s.description);
    }

    /* Dipakai bersama oleh detail Layanan & Fasilitas karena strukturnya sama
       (kategori, nama, deskripsi). */
    function openSimpleDetailModal(category, name, description) {
      const root = document.getElementById('modal-root');
      root.innerHTML = `
  <div class="modal-overlay" onclick="if(event.target===this) closeModal()">
    <div class="modal village-modal">
      <div class="vm-head">
        <div>
          <span class="program-cat">${escapeHtml(category)}</span>
          <h3 style="margin-top:6px;">${escapeHtml(name)}</h3>
        </div>
        <button class="vm-close" onclick="closeModal()" aria-label="Tutup">✕</button>
      </div>
      <div class="vm-body">
        <p class="vm-body-text">${escapeHtml(description)}</p>
      </div>
    </div>
  </div>`;
    }

    function renderClusters(list) {
      const el = document.getElementById('cluster-grid');
      __clustersActive = list;
      el.innerHTML = list.map((c, i) => `
    <div class="cluster-card" onclick="openClusterDetail(${i})" role="button" tabindex="0" onkeypress="if(event.key==='Enter') openClusterDetail(${i})">
      <div class="num">Klaster ${c.id}</div>
      <h4>${escapeHtml(c.title)}</h4>
      <p>${escapeHtml(c.description)}</p>
      <ul>${(c.items || []).map(i => `<li>${escapeHtml(i)}</li>`).join('')}</ul>
    </div>`).join('');
    }

    function openClusterDetail(i) {
      const c = __clustersActive[i];
      if (!c) return;
      const root = document.getElementById('modal-root');
      const itemsHtml = (c.items || []).map(it => `<li>${escapeHtml(it)}</li>`).join('');
      root.innerHTML = `
  <div class="modal-overlay" onclick="if(event.target===this) closeModal()">
    <div class="modal village-modal">
      <div class="vm-head">
        <div>
          <span class="program-cat">Klaster ${escapeHtml(String(c.id))}</span>
          <h3 style="margin-top:6px;">${escapeHtml(c.title)}</h3>
        </div>
        <button class="vm-close" onclick="closeModal()" aria-label="Tutup">✕</button>
      </div>
      <div class="vm-body">
        <p class="vm-body-text">${escapeHtml(c.description)}</p>
        ${itemsHtml ? `<ul class="vm-body-list">${itemsHtml}</ul>` : ''}
      </div>
    </div>
  </div>`;
    }

    function renderStaffNeeds(list) {
      const el = document.getElementById('staff-list');
      const active = list.filter(s => s.is_active);
      __staffActive = active;
      el.innerHTML = active.length ? active.map((s, i) => `
    <div class="staff-card" onclick="openStaffDetail(${i})" role="button" tabindex="0" onkeypress="if(event.key==='Enter') openStaffDetail(${i})">
      <div>
        <div class="dname">${escapeHtml(s.posisi)}</div>
        <div class="dsub">${escapeHtml(s.kualifikasi)}${s.keterangan ? (' · ' + escapeHtml(s.keterangan)) : ''}</div>
      </div>
      <div class="dtime">${s.jumlah} orang<br>Batas: ${s.batas_lamar || '-'}</div>
    </div>`).join('') : '<p style="color:var(--muted);">Belum ada informasi kebutuhan tenaga kesehatan.</p>';
    }

    function openStaffDetail(i) {
      const s = __staffActive[i];
      if (!s) return;
      const root = document.getElementById('modal-root');
      const hasForm = isSafeFileSrc(s.form_file);
      root.innerHTML = `
  <div class="modal-overlay" onclick="if(event.target===this) closeModal()">
    <div class="modal village-modal">
      <div class="vm-head">
        <div>
          <span class="program-cat">Kebutuhan Tenaga Kesehatan &amp; Tenaga Administrasi</span>
          <h3 style="margin-top:6px;">${escapeHtml(s.posisi)}</h3>
        </div>
        <button class="vm-close" onclick="closeModal()" aria-label="Tutup">✕</button>
      </div>
      <div class="vm-body">
        <p class="vm-body-text">${escapeHtml(s.kualifikasi)}</p>
        ${s.keterangan ? `<p class="vm-body-text">${escapeHtml(s.keterangan)}</p>` : ''}
        <div class="vm-meta-row"><span class="k">Jumlah dibutuhkan</span><span class="v">${escapeHtml(String(s.jumlah))} orang</span></div>
        <div class="vm-meta-row"><span class="k">Batas pendaftaran</span><span class="v">${escapeHtml(s.batas_lamar || '-')}</span></div>
        <div class="modal-actions" style="justify-content:flex-start;flex-wrap:wrap;margin-top:18px;">
          ${hasForm
          ? `<a class="btn btn-outline btn-sm" href="${s.form_file}" download="${escapeHtml(s.form_file_name || (s.posisi + '-formulir'))}">⬇ Unduh Formulir Pendaftaran</a>`
          : `<span style="font-size:12px;color:var(--muted);">Formulir belum diunggah admin.</span>`}
          <button class="btn btn-primary btn-sm" onclick="openApplicationUploadModal('${s.id}')">📤 Upload Berkas untuk Posisi Ini</button>
        </div>
      </div>
    </div>
  </div>`;
    }

    /* ---- Upload Berkas Pendaftaran (publik) — dilindungi kode Google Authenticator ---- */
    function openApplicationUploadModal(preselectId) {
      const root = document.getElementById('modal-root');
      const options = __staffActive.map(s => `<option value="${s.id}" ${s.id === preselectId ? 'selected' : ''}>${escapeHtml(s.posisi)}</option>`).join('');
      root.innerHTML = `
  <div class="modal-overlay" onclick="if(event.target===this) closeModal()">
    <div class="modal village-modal">
      <div class="vm-head">
        <div>
          <span class="program-cat">Pendaftaran</span>
          <h3 style="margin-top:6px;">Upload Berkas Pendaftaran</h3>
        </div>
        <button class="vm-close" onclick="closeModal()" aria-label="Tutup">✕</button>
      </div>
      <div class="vm-body">
        <p class="vm-body-text" style="margin-bottom:14px;">Isi data berikut, lampirkan berkas lamaran, lalu masukkan kode verifikasi 6 digit yang diberikan petugas Puskesmas (dihasilkan lewat Google Authenticator).</p>
        <div class="field"><label>Posisi yang Dilamar</label>
          <select id="ap-posisi">${options || '<option value="">Belum ada lowongan aktif</option>'}</select>
        </div>
        <div class="field"><label>Nama Lengkap</label><input id="ap-nama" placeholder="Nama sesuai KTP"></div>
        <div class="field"><label>No. HP / WhatsApp</label><input id="ap-kontak" placeholder="08xxxxxxxxxx"></div>
        <div class="field"><label>Berkas Lamaran (PDF/JPG/PNG, maks 5MB)</label>
          <input type="file" id="ap-file" accept=".pdf,.jpg,.jpeg,.png">
        </div>
        <div class="field"><label>Kode Verifikasi (Google Authenticator)</label>
          <input id="ap-otp" inputmode="numeric" maxlength="6" placeholder="6 digit">
        </div>
        <div class="modal-actions">
          <button class="btn btn-outline" onclick="closeModal()">Batal</button>
          <button class="btn btn-primary" id="ap-submit-btn" onclick="submitApplicationUpload()">Kirim Berkas</button>
        </div>
      </div>
    </div>
  </div>`;
    }

    async function submitApplicationUpload() {
      const posisiId = document.getElementById('ap-posisi').value;
      const nama = document.getElementById('ap-nama').value.trim();
      const kontak = document.getElementById('ap-kontak').value.trim();
      const fileInput = document.getElementById('ap-file');
      const otp = document.getElementById('ap-otp').value.trim();
      const btn = document.getElementById('ap-submit-btn');

      if (!posisiId) { toast('Tidak ada lowongan aktif untuk dilamar', true); return; }
      if (!nama) { toast('Nama lengkap wajib diisi', true); return; }
      if (!kontak) { toast('No. HP/WhatsApp wajib diisi', true); return; }
      if (!fileInput.files[0]) { toast('Berkas lamaran wajib dilampirkan', true); return; }
      if (!/^\d{6}$/.test(otp)) { toast('Kode verifikasi harus 6 digit angka', true); return; }

      btn.disabled = true; btn.textContent = 'Memverifikasi...';
      try {
        const settings = await DataLayer.getSettings();
        const secret = settings.totp_secret;
        if (!secret) { toast('Kode verifikasi belum diaktifkan oleh admin. Hubungi petugas Puskesmas.', true); return; }

        const valid = await totpVerify(secret, otp);
        if (!valid) { toast('Kode verifikasi salah atau sudah kedaluwarsa', true); return; }

        btn.textContent = 'Mengunggah...';
        const posisi = (__staffActive.find(s => s.id === posisiId) || {}).posisi || '';
        const fileDataUrl = await fileToDataUrl(fileInput.files[0], 5, ['.pdf', '.jpg', '.jpeg', '.png']);

        await DataLayer.insertApplication({
          staff_need_id: posisiId,
          posisi,
          nama,
          kontak,
          file_name: fileInput.files[0].name,
          file_data: fileDataUrl,
          created_at: new Date().toISOString()
        });

        closeModal();
        toast('Berkas pendaftaran berhasil dikirim');
      } catch (e) {
        toast('Gagal mengirim: ' + errMsg(e), true);
      } finally {
        if (btn) { btn.disabled = false; btn.textContent = 'Kirim Berkas'; }
      }
    }

    function renderDoctorSchedule(list) {
      const el = document.getElementById('doctor-list');
      const active = list.filter(s => s.is_active);
      __doctorsActive = active;
      el.innerHTML = active.length ? active.map((s, i) => `
    <div class="doctor-card" onclick="openDoctorDetail(${i})" role="button" tabindex="0" onkeypress="if(event.key==='Enter') openDoctorDetail(${i})">
      <div>
        <div class="dname">${escapeHtml(s.nama_dokter)}</div>
        <div class="dspec">${escapeHtml(s.spesialisasi)} · ${escapeHtml(s.lokasi || '')}</div>
        ${s.keterangan ? `<div class="dsub">${escapeHtml(s.keterangan)}</div>` : ''}
      </div>
      <div class="dtime">${escapeHtml(s.hari)}<br>${escapeHtml(s.jam)}</div>
    </div>`).join('') : '<p style="color:var(--muted);">Belum ada jadwal dokter spesialis.</p>';
    }

    function openDoctorDetail(i) {
      const s = __doctorsActive[i];
      if (!s) return;
      const root = document.getElementById('modal-root');
      root.innerHTML = `
  <div class="modal-overlay" onclick="if(event.target===this) closeModal()">
    <div class="modal village-modal">
      <div class="vm-head">
        <div>
          <span class="program-cat">${escapeHtml(s.spesialisasi)}</span>
          <h3 style="margin-top:6px;">${escapeHtml(s.nama_dokter)}</h3>
        </div>
        <button class="vm-close" onclick="closeModal()" aria-label="Tutup">✕</button>
      </div>
      <div class="vm-body">
        ${s.keterangan ? `<p class="vm-body-text">${escapeHtml(s.keterangan)}</p>` : ''}
        <div class="vm-meta-row"><span class="k">Lokasi</span><span class="v">${escapeHtml(s.lokasi || '-')}</span></div>
        <div class="vm-meta-row"><span class="k">Hari praktik</span><span class="v">${escapeHtml(s.hari)}</span></div>
        <div class="vm-meta-row"><span class="k">Jam praktik</span><span class="v">${escapeHtml(s.jam)}</span></div>
      </div>
    </div>
  </div>`;
    }

    /* Deteksi jenis tujuan hotline: nomor telepon, URL (web/IG/aplikasi lain), WhatsApp, atau email,
       lalu tentukan href yang tepat supaya klik langsung mengarah ke tujuan yang benar. */
    function getHotlineTarget(nomorRaw) {
      const nomor = (nomorRaw || '').trim();
      const isWhatsApp = /^(https?:\/\/(wa\.me|api\.whatsapp\.com)\/|wa\.me\/|whatsapp:)/i.test(nomor);
      const isUrl = /^(https?:\/\/|www\.)/i.test(nomor);
      const isEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(nomor);

      if (isWhatsApp) {
        const href = /^https?:\/\//i.test(nomor) ? nomor : ('https://' + nomor.replace(/^whatsapp:/i, 'wa.me/'));
        return { href, external: true, icon: '💬' };
      }

      if (isUrl) {
        const href = /^www\./i.test(nomor) ? ('https://' + nomor) : nomor;
        const appScheme = getAppDeepLink(href);
        return { href, external: true, icon: appScheme ? appScheme.icon : '🔗', appScheme: appScheme ? appScheme.scheme : null };
      }

      if (isEmail) {
        return { href: 'mailto:' + nomor, external: false, icon: '✉️' };
      }

      const dialNumber = nomor.replace(/[^\d+]/g, '');
      return { href: 'tel:' + dialNumber, external: false, icon: '📞' };
    }

    /* Cocokkan URL dengan pola platform yang dikenal supaya bisa dibuka langsung
       lewat aplikasinya (bukan cuma browser), dengan fallback otomatis ke web
       kalau aplikasinya tidak terpasang di perangkat. */
    function getAppDeepLink(url) {
      try {
        const u = new URL(url);
        const host = u.hostname.replace(/^www\./i, '').toLowerCase();
        const path = u.pathname.replace(/^\/+|\/+$/g, '');
        const firstSegment = path.split('/')[0] || '';

        if (host === 'instagram.com' && firstSegment) {
          return { icon: '📷', scheme: `instagram://user?username=${encodeURIComponent(firstSegment)}` };
        }
        if ((host === 'facebook.com' || host === 'fb.com') && path) {
          return { icon: '📘', scheme: `fb://facewebmodal/f?href=${encodeURIComponent(url)}` };
        }
        if (host === 'tiktok.com' && firstSegment.startsWith('@')) {
          return { icon: '🎵', scheme: `snssdk1233://user/profile/${encodeURIComponent(firstSegment.slice(1))}` };
        }
        if (host === 't.me' && firstSegment) {
          return { icon: '✈️', scheme: `tg://resolve?domain=${encodeURIComponent(firstSegment)}` };
        }
        if ((host === 'youtube.com' || host === 'youtu.be') && path) {
          return { icon: '▶️', scheme: `vnd.youtube://${path}` };
        }
        if (host === 'x.com' || host === 'twitter.com') {
          return { icon: '🐦', scheme: `twitter://user?screen_name=${encodeURIComponent(firstSegment)}` };
        }
      } catch (e) { /* URL tidak valid untuk parsing, biarkan jadi link web biasa */ }
      return null;
    }

    /* Coba buka aplikasi native lewat custom URL scheme; kalau dalam waktu singkat
       halaman masih terlihat (artinya aplikasi tidak terpasang), fallback ke link web. */
    function openHotlineLink(event, webUrl, appScheme) {
      if (!appScheme) return; // biarkan browser buka link web seperti biasa
      event.preventDefault();
      const fallbackTimer = setTimeout(() => {
        if (!document.hidden) window.open(webUrl, '_blank', 'noopener,noreferrer');
      }, 1200);
      window.addEventListener('visibilitychange', function onHide() {
        if (document.hidden) { clearTimeout(fallbackTimer); window.removeEventListener('visibilitychange', onHide); }
      });
      window.location.href = appScheme;
    }

    function renderHotline(list) {
      const el = document.getElementById('hotline-list');
      const active = list.filter(h => h.is_active);
      el.innerHTML = active.length ? active.map(h => {
        const target = getHotlineTarget(h.nomor);
        const extraAttrs = target.external ? ' target="_blank" rel="noopener noreferrer"' : '';
        const clickAttr = target.appScheme ? ` onclick="openHotlineLink(event, this.dataset.web, this.dataset.app)" data-web="${escapeHtml(target.href)}" data-app="${escapeHtml(target.appScheme)}"` : '';
        return `
    <a class="hotline-card" href="${escapeHtml(target.href)}"${extraAttrs}${clickAttr} aria-label="Hubungi ${escapeHtml(h.nama)}: ${escapeHtml(h.nomor)}">
      <div>
        <div class="dname">${escapeHtml(h.nama)}</div>
        ${h.keterangan ? `<div class="dsub">${escapeHtml(h.keterangan)}</div>` : ''}
      </div>
      <div class="dtime"><span class="hotline-call-btn">${target.icon}<span>${escapeHtml(h.nomor)}</span></span></div>
    </a>`;
      }).join('') : '<p style="color:var(--muted);">Belum ada informasi hotline.</p>';
    }

    function renderBreakingNews(list) {
      const el = document.getElementById('breaking-grid');
      /* BUGFIX: saring gambar dengan sumber tidak aman (mis. skema javascript:) sebelum dirender sebagai src */
      const active = list.filter(n => n.is_active && isSafeImageSrc(n.image));
      if (!active.length) { el.innerHTML = '<p style="color:var(--muted);">Belum ada poster/flyer yang ditambahkan.</p>'; return; }
      el.innerHTML = active.map(n => `
    <div class="breaking-card">
      <div class="bframe"><img src="${escapeHtml(n.image)}" alt="${escapeHtml(n.title || 'Poster')}"></div>
      ${n.title ? `<div class="btitle">${escapeHtml(n.title)}</div>` : ''}
    </div>`).join('');
    }

    /* =====================================================================
       AUTH
       ===================================================================== */
    document.getElementById('login-form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const email = document.getElementById('login-email').value.trim();
      const pass = document.getElementById('login-password').value;
      const msg = document.getElementById('login-msg');
      msg.innerHTML = '';
      try {
        if (DEMO_MODE) {
          throw new Error('Supabase belum terhubung. Isi SUPABASE_URL dan SUPABASE_ANON_KEY pada bagian CONFIG di berkas ini terlebih dahulu.');
        } else {
          const { data, error } = await sb.auth.signInWithPassword({ email, password: pass });
          if (error) throw error;
          state.session = data.session;

          /* Cek status aktif/nonaktif akun. Jika akun ini belum punya baris di admin_profiles
             (akun lama, dibuat sebelum fitur ini ada), buatkan otomatis agar tidak terkunci. */
          try {
            const uid_ = data.session.user.id;
            const uemail = data.session.user.email;
            let profile = await DataLayer.getProfileById(uid_);
            if (!profile) {
              await DataLayer.ensureOwnProfile(uid_, uemail);
              profile = await DataLayer.getProfileById(uid_);
            }
            if (profile && profile.is_active === false) {
              await sb.auth.signOut();
              state.session = null;
              msg.innerHTML = `<div class="form-msg err">Akun ini telah dinonaktifkan. Hubungi admin lain untuk mengaktifkannya kembali.</div>`;
              return;
            }
            state.myProfile = profile;
          } catch (profileErr) {
            /* Jangan blokir login hanya karena tabel admin_profiles belum dibuat di Supabase;
               fitur role/nonaktif memang butuh migrasi SQL tambahan (lihat komentar CONFIG). */
            console.warn('Gagal memuat admin_profiles (mungkin tabel belum dibuat):', profileErr.message);
          }

          location.hash = '#/admin';
        }
      } catch (err) {
        msg.innerHTML = `<div class="form-msg err">${escapeHtml(errMsg(err))}</div>`;
      }
    });

    document.getElementById('btn-logout').addEventListener('click', async () => {
      if (DEMO_MODE) { demoAuthed = false; } else { await sb.auth.signOut(); state.session = null; }
      location.hash = '#/';
    });

    async function initSession() {
      if (DEMO_MODE) return;
      try {
        const { data, error } = await sb.auth.getSession();
        if (error) throw error;
        state.session = data.session;
        sb.auth.onAuthStateChange((_evt, session) => { state.session = session; });
      } catch (e) {
        /* Jangan biarkan kegagalan mengambil sesi awal (mis. proyek Supabase belum
           dikonfigurasi dengan benar, atau jaringan diblokir) menghentikan aplikasi
           dengan error yang tidak tertangkap. Anggap saja belum ada sesi aktif. */
        console.warn('Gagal memuat sesi awal:', errMsg(e));
        state.session = null;
      }
    }

    /* Muat ulang profil (peran + status aktif) akun yang sedang login setiap kali panel admin
       dibuka, bukan hanya sesaat setelah submit form login. Ini juga menutup celah: akun yang
       dinonaktifkan admin lain saat masih dalam sesi terbuka (tab lama) akan langsung ditolak
       begitu me-reload/navigasi ulang, bukan tetap bisa memakai panel admin. Mengembalikan
       false jika akun ternyata sudah dinonaktifkan (sesi akan otomatis diputus). */
    async function ensureMyProfileLoaded() {
      if (DEMO_MODE || !state.session || !state.session.user) return true;
      const uid_ = state.session.user.id;
      const uemail = state.session.user.email;
      if (state.myProfile && state.myProfile.id === uid_) {
        /* Sudah ada dari sesi login sebelumnya; tetap segarkan status aktifnya saja. */
      }
      try {
        let profile = await DataLayer.getProfileById(uid_);
        if (!profile) {
          await DataLayer.ensureOwnProfile(uid_, uemail);
          profile = await DataLayer.getProfileById(uid_);
        }
        if (profile && profile.is_active === false) {
          await sb.auth.signOut();
          state.session = null;
          state.myProfile = null;
          toast('Akun ini telah dinonaktifkan. Hubungi admin lain untuk mengaktifkannya kembali.', true);
          return false;
        }
        state.myProfile = profile;
      } catch (profileErr) {
        /* Jangan blokir akses hanya karena tabel admin_profiles belum dibuat di Supabase. */
        console.warn('Gagal memuat admin_profiles (mungkin tabel belum dibuat):', profileErr.message);
      }
      return true;
    }

    /* =====================================================================
       ADMIN VIEW
       ===================================================================== */
    document.querySelectorAll('.nav-link[data-tab]').forEach(btn => {
      btn.addEventListener('click', () => location.hash = '#/admin/' + btn.dataset.tab);
    });

    function setAdminTab(tab) {
      document.querySelectorAll('.nav-link[data-tab]').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
      document.querySelectorAll('.admin-tab').forEach(t => t.classList.add('hidden'));
      const titles = {
        ringkasan: ['Ringkasan', 'Ikhtisar data papan informasi'],
        program: ['Capaian Layanan', 'Kelola data capaian program kesehatan per bulan (Januari–Desember)'],
        pengumuman: ['Papan Informasi', 'Kelola teks berjalan & pengumuman di halaman publik'],
        layanan: ['Informasi Layanan', 'Kelola daftar jenis layanan kesehatan'],
        fasilitas: ['Fasilitas Kesehatan', 'Kelola daftar fasilitas layanan kesehatan'],
        klaster: ['Klaster Layanan', 'Tambah, ubah, atau hapus klaster beserta deskripsi dan daftar layanannya'],
        sdm: ['SDM & Jadwal Dokter', 'Kebutuhan tenaga kesehatan/administrasi, jadwal dokter, formulir & berkas pendaftaran'],
        hotline: ['Hotline & Kontak', 'Kelola nomor hotline layanan dan kontak darurat'],
        breaking: ['Breaking News', 'Kelola poster dan flyer informasi terbaru'],
        jam: ['Jam Layanan', 'Atur jadwal layanan per hari'],
        tampilan: ['Pengaturan Tampilan', 'Sesuaikan profil, warna, kecepatan teks berjalan, dan tata letak'],
        akun: ['Akun', 'Kelola email/kata sandi akun ini, atau tambahkan akun admin baru']
      };
      const t = titles[tab] || titles.ringkasan;
      document.getElementById('admin-title').textContent = t[0];
      document.getElementById('admin-sub').textContent = t[1];
      const panel = document.getElementById('tab-' + tab);
      if (panel) panel.classList.remove('hidden');
      document.getElementById('admin-sidebar').classList.remove('open');
    }

    async function loadAdminData() {
      try {
        const [settings, hours, ann, programs, services, facilities, clusters, staffNeeds, doctorSchedule, hotline, breakingNews, applications] = await Promise.all([
          DataLayer.getSettings(), DataLayer.getHours(), DataLayer.getAnnouncements(true), DataLayer.getPrograms(),
          DataLayer.getServices(), DataLayer.getFacilities(), DataLayer.getClusters(), DataLayer.getStaffNeeds(), DataLayer.getDoctorSchedule(), DataLayer.getHotline(), DataLayer.getBreakingNews(), DataLayer.getApplications()
        ]);
        state.settings = settings; state.hours = hours; state.announcements = ann; state.programs = programs;
        state.services = services; state.facilities = facilities; state.clusters = clusters;
        state.staffNeeds = staffNeeds; state.doctorSchedule = doctorSchedule; state.hotline = hotline; state.breakingNews = breakingNews;
        state.applications = applications;

        renderRingkasan();
        renderProgramTable();
        renderAnnounceTable();
        renderTickerPreview();
        renderServicesTable();
        renderFacilitiesTable();
        renderClusterForms();
        renderStaffTable();
        renderDoctorTable();
        renderHotlineTable();
        renderBreakingNewsTable();
        renderHoursTable();
        renderTampilanForm();
        renderAkunForm();
        await renderPenggunaTable();
        renderAuditLog();
        renderTotpPanel();
        renderApplicationsTable();
      } catch (e) { console.error(e); toast('Gagal memuat data admin: ' + errMsg(e), true); }
    }

    /* ---- Ringkasan ---- */
    function renderRingkasan() {
      const series = groupProgramSeries(state.programs);
      document.getElementById('stat-programs').textContent = series.length;

      const latestPercents = series.map(s => {
        const months = Object.keys(s.monthly).map(Number);
        const lm = months.length ? Math.max(...months) : null;
        return lm ? pct(s.monthly[lm], s.target) : 0;
      });
      const avg = latestPercents.length ? Math.round(latestPercents.reduce((a, b) => a + b, 0) / latestPercents.length) : 0;
      document.getElementById('stat-avg').textContent = avg + '%';
      document.getElementById('stat-announce').textContent = state.announcements.filter(a => a.is_active).length;
      document.getElementById('stat-days').textContent = state.hours.filter(h => !h.is_closed).length + '/7';

      const withPct = series.map((s, i) => ({ s, p: latestPercents[i] })).sort((a, b) => b.p - a.p);
      /* BUGFIX: sebelumnya slice(0,3) dan slice(-3) bisa tumpang tindih ketika total program <= 6,
         sehingga program yang sama muncul dobel di "Tertinggi" dan "Perlu Perhatian". */
      const topCount = Math.min(3, withPct.length);
      const top = withPct.slice(0, topCount);
      const bottomCount = Math.min(3, withPct.length - topCount);
      const bottom = bottomCount > 0 ? withPct.slice(withPct.length - bottomCount).reverse() : [];
      const row = (x, label) => `<div style="display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid var(--border);font-size:13px;">
    <span>${label} ${escapeHtml(x.s.name)}</span><b style="font-family:var(--font-mono);">${x.p}%</b></div>`;
      document.getElementById('ringkasan-list').innerHTML =
        (top.length ? '<div style="font-weight:700;font-size:12.5px;color:var(--muted);margin-bottom:4px;">TERTINGGI</div>' + top.map(x => row(x, '▲')).join('') : '') +
        (bottom.length ? '<div style="font-weight:700;font-size:12.5px;color:var(--muted);margin:14px 0 4px;">PERLU PERHATIAN</div>' + bottom.map(x => row(x, '▼')).join('') : '<p class="desc">Belum ada data.</p>');
    }

    /* ---- Program & Capaian (bulanan) ---- */
    function renderProgramTable() {
      const catSel = document.getElementById('program-cat-filter');
      const cats = [...new Set(state.programs.map(p => p.category))];
      const curVal = catSel.value;
      catSel.innerHTML = '<option value="">Semua Kategori</option>' + cats.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
      catSel.value = curVal;

      const q = (document.getElementById('program-search').value || '').toLowerCase();
      const catF = catSel.value;
      const list = state.programs
        .filter(p => (!catF || p.category === catF) && p.name.toLowerCase().includes(q))
        .sort((a, b) => a.name.localeCompare(b.name) || String(b.period).localeCompare(String(a.period)) || a.month - b.month || (a.desa || DESA_ALL).localeCompare(b.desa || DESA_ALL));

      document.getElementById('program-tbody').innerHTML = list.map(p => `
    <tr>
      <td><b>${escapeHtml(p.name)}</b></td>
      <td><span class="badge">${escapeHtml(p.category)}</span></td>
      <td><span class="badge${isAggregateDesa(p.desa) ? '' : ' off'}">${escapeHtml(p.desa || DESA_ALL)}</span></td>
      <td>${escapeHtml(p.period)}</td>
      <td>${MONTH_SHORT[p.month] || '-'}</td>
      <td>${p.target}${p.unit}</td>
      <td>${p.achieved}${p.unit}</td>
      <td><b style="color:${achievementColor(pct(p.achieved, p.target), (state.settings.theme && state.settings.theme.chart), 'var(--primary-dark)')}">${pct(p.achieved, p.target)}%</b></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openProgramModal('${p.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeProgram('${p.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="9" style="color:var(--muted);text-align:center;padding:20px;">Tidak ada data</td></tr>';
    }
    document.getElementById('program-search').addEventListener('input', renderProgramTable);
    document.getElementById('program-cat-filter').addEventListener('change', renderProgramTable);

    function openProgramModal(id) {
      const p = id ? state.programs.find(x => x.id === id) : { name: '', category: '', period: new Date().getFullYear().toString(), month: new Date().getMonth() + 1, target: 100, achieved: 0, unit: '%', desa: DESA_ALL };
      const root = document.getElementById('modal-root');
      const monthOptions = MONTH_NAMES.map((m, i) => i === 0 ? '' : `<option value="${i}" ${p.month === i ? 'selected' : ''}>${m}</option>`).join('');
      const curDesa = p.desa || DESA_ALL;
      const desaOptions = [DESA_ALL, ...DESA_LIST].map(d => `<option value="${escapeHtml(d)}" ${curDesa === d ? 'selected' : ''}>${escapeHtml(d)}</option>`).join('');
      /* Kategori dipilih lewat dropdown (bukan ketik bebas) supaya kategori yang
         sudah ada tidak tercatat dobel gara-gara beda kapitalisasi/spasi (mis.
         "Kesling" vs "kesling "). Pilihan terakhir "+ Tambah kategori baru…"
         membuka kotak isian baru — kategori baru itu otomatis disamakan ejaannya
         dengan kategori yang sudah ada bila cocok (case-insensitive). */
      const existingCats = [...new Set((state.programs || []).map(x => x.category).filter(Boolean))].sort((a, b) => a.localeCompare(b));
      const curCatMatch = existingCats.find(c => c.toLowerCase() === (p.category || '').toLowerCase());
      const catOptions = existingCats.map(c => `<option value="${escapeHtml(c)}" ${curCatMatch === c ? 'selected' : ''}>${escapeHtml(c)}</option>`).join('')
        + `<option value="__new__" ${(!curCatMatch) ? 'selected' : ''}>+ Tambah kategori baru…</option>`;
      root.innerHTML = `
  <div class="modal-overlay">
    <div class="modal">
      <h3>${id ? 'Ubah Data Capaian' : 'Tambah Data Capaian Bulan'}</h3>
      <div class="field"><label>Nama Program</label><input id="m-name" value="${escapeHtml(p.name)}"></div>
      <div class="field-row">
        <div class="field"><label>Kategori</label>
          <select id="m-category" onchange="document.getElementById('m-category-new-wrap').style.display=(this.value==='__new__')?'block':'none';">${catOptions}</select>
        </div>
        <div class="field"><label>Periode (Tahun)</label><input id="m-period" value="${escapeHtml(p.period)}"></div>
      </div>
      <div id="m-category-new-wrap" class="field" style="display:${curCatMatch ? 'none' : 'block'};margin-top:-6px;">
        <label>Nama Kategori Baru</label>
        <input id="m-category-new" value="${curCatMatch ? '' : escapeHtml(p.category || '')}" placeholder="mis. Imunisasi, KIA, Gizi">
      </div>
      <div class="field-row">
        <div class="field"><label>Bulan</label>
          <select id="m-month">${monthOptions}</select>
        </div>
        <div class="field"><label>Desa</label>
          <select id="m-desa">${desaOptions}</select>
        </div>
      </div>
      <p class="desc" style="margin:-4px 0 4px;">Pilih <b>${escapeHtml(DESA_ALL)}</b> untuk data total Puskesmas (dipakai pada grafik utama), atau pilih salah satu desa untuk mengisi rincian capaian desa tersebut yang tampil saat kartu Capaian Layanan diklik oleh publik. <b>Tidak wajib mengisi semua desa</b> — desa yang belum diisi akan otomatis direkap/diperkirakan sistem dari data total Puskesmas di atas.</p>
      <div class="field-row">
        <div class="field"><label>Target</label><input id="m-target" type="number" value="${p.target}"></div>
        <div class="field"><label>Capaian</label><input id="m-achieved" type="number" value="${p.achieved}"></div>
        <div class="field"><label>Satuan</label><input id="m-unit" value="${escapeHtml(p.unit)}"></div>
      </div>
      <div class="modal-actions">
        <button class="btn btn-outline" onclick="closeModal()">Batal</button>
        <button class="btn btn-primary" onclick="saveProgramModal('${id || ''}')">Simpan</button>
      </div>
    </div>
  </div>`;
    }
    async function saveProgramModal(id) {
      const catSelVal = document.getElementById('m-category').value;
      const catRaw = catSelVal === '__new__' ? document.getElementById('m-category-new').value : catSelVal;
      const row = {
        id: id || undefined,
        name: document.getElementById('m-name').value.trim(),
        category: canonicalizeCategory(catRaw || 'Umum'),
        period: document.getElementById('m-period').value.trim(),
        month: parseInt(document.getElementById('m-month').value) || 1,
        desa: document.getElementById('m-desa').value || DESA_ALL,
        target: parseFloat(document.getElementById('m-target').value) || 0,
        achieved: parseFloat(document.getElementById('m-achieved').value) || 0,
        unit: document.getElementById('m-unit').value.trim() || '%'
      };
      if (!row.name) { toast('Nama program wajib diisi', true); return; }
      try { await DataLayer.upsertProgram(row); closeModal(); await loadAdminData(); toast('Data capaian tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeProgram(id) {
      if (!confirm('Hapus data ini?')) return;
      try { await DataLayer.deleteProgram(id); await loadAdminData(); toast('Data dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    function closeModal() { document.getElementById('modal-root').innerHTML = ''; }
    document.getElementById('btn-add-program').addEventListener('click', () => openProgramModal(null));

    /* Ekspor/impor Capaian Layanan — kolom & urutan mengikuti persis template
       Excel resmi (Report.xlsx): name, category, desa, period year, month,
       target, achieved, unit — supaya file yang diekspor bisa langsung
       diimpor kembali tanpa perlu dirapikan ulang. */
    const EXPORT_COLUMNS = ['name', 'category', 'desa', 'period year', 'month', 'target', 'achieved', 'unit'];
    function buildExportRows() {
      return state.programs.map(p => ({
        name: p.name,
        category: p.category,
        desa: p.desa || DESA_ALL,
        'period year': p.period,
        month: p.month,
        target: p.target,
        achieved: p.achieved,
        unit: p.unit
      }));
    }

    /* Ekspor Excel (.xlsx) — format utama, sama seperti file yang dikirim pengguna */
    document.getElementById('btn-export-xlsx').addEventListener('click', () => {
      try {
        const rows = buildExportRows();
        const ws = XLSX.utils.json_to_sheet(rows, { header: EXPORT_COLUMNS });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, 'capaian-layanan-puskesmas-puger.xlsx');
        toast('Excel diekspor');
      } catch (err) { toast('Gagal mengekspor Excel: ' + errMsg(err), true); }
    });

    /* Ekspor CSV — kolom sama persis dengan template Excel, untuk yang lebih suka CSV */
    document.getElementById('btn-export-csv').addEventListener('click', () => {
      const rows = buildExportRows();
      const csv = Papa.unparse(rows, { columns: EXPORT_COLUMNS });
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'capaian-layanan-puskesmas-puger.csv';
      a.click();
      toast('CSV diekspor');
    });

    /* Impor Excel (.xlsx/.xls) atau CSV — header dicocokkan fleksibel via
       normalizedKeyMap()/pickField() sehingga "period year", "Periode", atau
       "period" semuanya terbaca sebagai kolom periode yang sama. */
    async function handleImportedRows(rawRows, fileInput) {
      try {
        const seenCat = new Map();
        const rows = rawRows.map(r => {
          const map = normalizedKeyMap(r);
          return {
            name: String(pickField(map, 'name', 'nama') || '').trim(),
            category: canonicalizeCategory(pickField(map, 'category', 'kategori') || 'Umum', seenCat),
            desa: String(pickField(map, 'desa') || '').trim() || DESA_ALL,
            period: String(pickField(map, 'period year', 'periodyear', 'period', 'periode', 'periodetahun') || new Date().getFullYear().toString()).trim(),
            month: parseMonth(pickField(map, 'month', 'bulan')),
            target: parseFloat(pickField(map, 'target')) || 0,
            achieved: parseFloat(pickField(map, 'achieved', 'capaian', 'realisasi')) || 0,
            unit: String(pickField(map, 'unit', 'satuan') || '%').trim()
          };
        }).filter(r => r.name);
        if (!rows.length) { toast('Berkas tidak berisi data valid', true); return; }
        await DataLayer.bulkImportPrograms(rows);
        await loadAdminData();
        toast(`${rows.length} baris berhasil diimpor`);
      } catch (err) { toast('Gagal mengimpor data: ' + errMsg(err), true); }
      finally { if (fileInput) fileInput.value = ''; }
    }
    document.getElementById('csv-input').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      if (ext === 'csv') {
        Papa.parse(file, {
          header: true, skipEmptyLines: true,
          complete: async (res) => { await handleImportedRows(res.data, e.target); },
          error: (err) => { toast('Gagal membaca CSV: ' + errMsg(err), true); e.target.value = ''; }
        });
      } else {
        const reader = new FileReader();
        reader.onload = async (ev) => {
          try {
            const data = new Uint8Array(ev.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            await handleImportedRows(json, e.target);
          } catch (err) { toast('Gagal membaca Excel: ' + errMsg(err), true); e.target.value = ''; }
        };
        reader.onerror = () => { toast('Gagal membaca berkas Excel', true); e.target.value = ''; };
        reader.readAsArrayBuffer(file);
      }
    });

    /* ---- Pengumuman / Teks Berjalan ---- */
    function renderAnnounceTable() {
      document.getElementById('announce-tbody').innerHTML = state.announcements.map(a => `
    <tr>
      <td><b>${escapeHtml(a.title)}</b><div style="color:var(--muted);font-size:11.5px;margin-top:2px;">${escapeHtml((a.content || '').slice(0, 70))}</div></td>
      <td>${a.priority}</td>
      <td><span class="badge ${a.is_active ? '' : 'off'}">${a.is_active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openAnnounceModal('${a.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeAnnounce('${a.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:20px;">Belum ada pengumuman</td></tr>';
    }
    function openAnnounceModal(id) {
      const a = id ? state.announcements.find(x => x.id === id) : { title: '', content: '', priority: 0, is_active: true };
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay">
    <div class="modal">
      <h3>${id ? 'Ubah Pengumuman' : 'Tambah Pengumuman'}</h3>
      <div class="field"><label>Judul</label><input id="m-title" value="${escapeHtml(a.title)}"></div>
      <div class="field"><label>Isi</label><textarea id="m-content" rows="3">${escapeHtml(a.content || '')}</textarea></div>
      <div class="field-row">
        <div class="field"><label>Prioritas (0-9)</label><input id="m-priority" type="number" min="0" max="9" value="${a.priority}"></div>
        <div class="field"><label>Status</label>
          <select id="m-active"><option value="true" ${a.is_active ? 'selected' : ''}>Aktif</option><option value="false" ${!a.is_active ? 'selected' : ''}>Nonaktif</option></select>
        </div>
      </div>
      <div class="modal-actions">
        <button class="btn btn-outline" onclick="closeModal()">Batal</button>
        <button class="btn btn-primary" onclick="saveAnnounceModal('${id || ''}')">Simpan</button>
      </div>
    </div>
  </div>`;
    }
    async function saveAnnounceModal(id) {
      const row = {
        id: id || undefined,
        title: document.getElementById('m-title').value.trim(),
        content: document.getElementById('m-content').value.trim(),
        priority: parseInt(document.getElementById('m-priority').value) || 0,
        is_active: document.getElementById('m-active').value === 'true'
      };
      if (!row.title) { toast('Judul wajib diisi', true); return; }
      try { await DataLayer.upsertAnnouncement(row); closeModal(); await loadAdminData(); toast('Pengumuman tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeAnnounce(id) {
      if (!confirm('Hapus pengumuman ini?')) return;
      try { await DataLayer.deleteAnnouncement(id); await loadAdminData(); toast('Pengumuman dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-announce').addEventListener('click', () => openAnnounceModal(null));

    function renderTickerPreview() {
      const el = document.getElementById('ticker-live-preview');
      const active = state.announcements.filter(a => a.is_active);
      el.textContent = active.length ? active.map(a => a.title).join('   •   ') : 'Belum ada teks berjalan aktif.';
    }
    document.getElementById('btn-quick-ticker').addEventListener('click', async () => {
      const input = document.getElementById('quick-ticker-input');
      const text = input.value.trim();
      if (!text) { toast('Isi teks terlebih dahulu', true); return; }
      try {
        await DataLayer.upsertAnnouncement({ title: text, content: '', priority: 0, is_active: true });
        input.value = '';
        await loadAdminData();
        toast('Teks berjalan ditambahkan');
      } catch (e) { toast('Gagal menambahkan: ' + errMsg(e), true); }
    });

    /* ---- Informasi Layanan ---- */
    function renderServicesTable() {
      document.getElementById('service-tbody').innerHTML = state.services.map(s => `
    <tr>
      <td><b>${escapeHtml(s.name)}</b></td>
      <td><span class="badge">${escapeHtml(s.category)}</span></td>
      <td style="max-width:260px;">${escapeHtml((s.description || '').slice(0, 80))}</td>
      <td><span class="badge ${s.is_active ? '' : 'off'}">${s.is_active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openServiceModal('${s.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeService('${s.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="5" style="color:var(--muted);text-align:center;padding:20px;">Belum ada data</td></tr>';
    }
    function openServiceModal(id) {
      const s = id ? state.services.find(x => x.id === id) : { name: '', category: '', description: '', is_active: true };
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay"><div class="modal">
    <h3>${id ? 'Ubah Layanan' : 'Tambah Layanan'}</h3>
    <div class="field"><label>Nama Layanan</label><input id="m-s-name" value="${escapeHtml(s.name)}"></div>
    <div class="field"><label>Kategori</label><input id="m-s-category" value="${escapeHtml(s.category)}" placeholder="mis. Rawat Jalan, Penunjang"></div>
    <div class="field"><label>Deskripsi</label><textarea id="m-s-desc" rows="3">${escapeHtml(s.description)}</textarea></div>
    <div class="field"><label>Status</label>
      <select id="m-s-active"><option value="true" ${s.is_active ? 'selected' : ''}>Aktif</option><option value="false" ${!s.is_active ? 'selected' : ''}>Nonaktif</option></select>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="saveServiceModal('${id || ''}')">Simpan</button>
    </div>
  </div></div>`;
    }
    async function saveServiceModal(id) {
      const row = {
        id: id || undefined,
        name: document.getElementById('m-s-name').value.trim(),
        category: document.getElementById('m-s-category').value.trim() || 'Umum',
        description: document.getElementById('m-s-desc').value.trim(),
        is_active: document.getElementById('m-s-active').value === 'true'
      };
      if (!row.name) { toast('Nama layanan wajib diisi', true); return; }
      try { await DataLayer.upsertService(row); closeModal(); await loadAdminData(); toast('Layanan tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeService(id) {
      if (!confirm('Hapus layanan ini?')) return;
      try { await DataLayer.deleteService(id); await loadAdminData(); toast('Layanan dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-service').addEventListener('click', () => openServiceModal(null));

    /* ---- Fasilitas Kesehatan ---- */
    function renderFacilitiesTable() {
      document.getElementById('facility-tbody').innerHTML = state.facilities.map(s => `
    <tr>
      <td><b>${escapeHtml(s.name)}</b></td>
      <td><span class="badge">${escapeHtml(s.category)}</span></td>
      <td style="max-width:260px;">${escapeHtml((s.description || '').slice(0, 80))}</td>
      <td><span class="badge ${s.is_active ? '' : 'off'}">${s.is_active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openFacilityModal('${s.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeFacility('${s.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="5" style="color:var(--muted);text-align:center;padding:20px;">Belum ada data</td></tr>';
    }
    function openFacilityModal(id) {
      const s = id ? state.facilities.find(x => x.id === id) : { name: '', category: '', description: '', is_active: true };
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay"><div class="modal">
    <h3>${id ? 'Ubah Fasilitas' : 'Tambah Fasilitas'}</h3>
    <div class="field"><label>Nama Fasilitas</label><input id="m-f-name" value="${escapeHtml(s.name)}"></div>
    <div class="field"><label>Kategori</label><input id="m-f-category" value="${escapeHtml(s.category)}" placeholder="mis. Fasilitas Gedung, Transportasi"></div>
    <div class="field"><label>Deskripsi</label><textarea id="m-f-desc" rows="3">${escapeHtml(s.description)}</textarea></div>
    <div class="field"><label>Status</label>
      <select id="m-f-active"><option value="true" ${s.is_active ? 'selected' : ''}>Aktif</option><option value="false" ${!s.is_active ? 'selected' : ''}>Nonaktif</option></select>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="saveFacilityModal('${id || ''}')">Simpan</button>
    </div>
  </div></div>`;
    }
    async function saveFacilityModal(id) {
      const row = {
        id: id || undefined,
        name: document.getElementById('m-f-name').value.trim(),
        category: document.getElementById('m-f-category').value.trim() || 'Umum',
        description: document.getElementById('m-f-desc').value.trim(),
        is_active: document.getElementById('m-f-active').value === 'true'
      };
      if (!row.name) { toast('Nama fasilitas wajib diisi', true); return; }
      try { await DataLayer.upsertFacility(row); closeModal(); await loadAdminData(); toast('Fasilitas tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeFacility(id) {
      if (!confirm('Hapus fasilitas ini?')) return;
      try { await DataLayer.deleteFacility(id); await loadAdminData(); toast('Fasilitas dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-facility').addEventListener('click', () => openFacilityModal(null));

    /* ---- Klaster Layanan ---- */
    function renderClusterForms() {
      const el = document.getElementById('cluster-admin-list');
      el.innerHTML = state.clusters.length ? state.clusters.map(c => `
    <div class="panel">
      <div class="toolbar">
        <div class="left"><h3 style="margin:0;">Klaster ${c.id}</h3></div>
        <button class="btn btn-danger btn-sm" onclick="removeCluster(${c.id})">Hapus Klaster</button>
      </div>
      <div class="field"><label>Judul Klaster</label><input id="cl-title-${c.id}" value="${escapeHtml(c.title)}"></div>
      <div class="field"><label>Deskripsi</label><textarea id="cl-desc-${c.id}" rows="2">${escapeHtml(c.description)}</textarea></div>
      <div class="field"><label>Daftar Layanan dalam Klaster (satu per baris)</label><textarea id="cl-items-${c.id}" rows="4">${escapeHtml((c.items || []).join('\n'))}</textarea></div>
      <button class="btn btn-primary btn-sm" onclick="saveCluster(${c.id})">Simpan Klaster ${c.id}</button>
    </div>`).join('') : '<div class="panel"><p class="desc" style="margin:0;">Belum ada klaster. Klik "+ Tambah Klaster" untuk membuat klaster baru.</p></div>';
    }
    async function saveCluster(id) {
      const row = {
        id,
        title: document.getElementById('cl-title-' + id).value.trim(),
        description: document.getElementById('cl-desc-' + id).value.trim(),
        items: document.getElementById('cl-items-' + id).value.split('\n').map(x => x.trim()).filter(Boolean)
      };
      if (!row.title) { toast('Judul klaster wajib diisi', true); return; }
      try { await DataLayer.updateCluster(row); await loadAdminData(); toast('Klaster ' + id + ' tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeCluster(id) {
      if (!confirm('Hapus klaster ini beserta seluruh isinya?')) return;
      try { await DataLayer.deleteCluster(id); await loadAdminData(); toast('Klaster dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    function openClusterModal() {
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay"><div class="modal">
    <h3>Tambah Klaster</h3>
    <div class="field"><label>Judul Klaster</label><input id="m-cl-title" placeholder="mis. Klaster 5 – Layanan Tambahan"></div>
    <div class="field"><label>Deskripsi</label><textarea id="m-cl-desc" rows="2"></textarea></div>
    <div class="field"><label>Daftar Layanan dalam Klaster (satu per baris)</label><textarea id="m-cl-items" rows="4"></textarea></div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="saveClusterModal()">Simpan</button>
    </div>
  </div></div>`;
    }
    async function saveClusterModal() {
      const row = {
        title: document.getElementById('m-cl-title').value.trim(),
        description: document.getElementById('m-cl-desc').value.trim(),
        items: document.getElementById('m-cl-items').value.split('\n').map(x => x.trim()).filter(Boolean)
      };
      if (!row.title) { toast('Judul klaster wajib diisi', true); return; }
      try { await DataLayer.upsertCluster(row); closeModal(); await loadAdminData(); toast('Klaster ditambahkan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-cluster').addEventListener('click', openClusterModal);

    /* ---- SDM: Kebutuhan Tenaga Kesehatan & Tenaga Administrasi ---- */
    /* Menyimpan sementara data URL formulir yang baru dipilih admin di modal tambah/ubah,
       sebelum tombol "Simpan" ditekan. undefined = tidak diubah, null = dihapus. */
    let pendingStaffFormFile, pendingStaffFormFileName;

    function renderStaffTable() {
      document.getElementById('staff-tbody').innerHTML = state.staffNeeds.map(s => `
    <tr>
      <td><b>${escapeHtml(s.posisi)}</b></td>
      <td>${escapeHtml(s.kualifikasi)}</td>
      <td>${s.jumlah}</td>
      <td>${s.batas_lamar || '-'}</td>
      <td>${isSafeFileSrc(s.form_file) ? `<span style="font-size:12px;color:var(--primary-dark);">📄 ${escapeHtml(s.form_file_name || 'formulir')}</span>` : '<span style="font-size:12px;color:var(--muted);">Belum ada</span>'}</td>
      <td><span class="badge ${s.is_active ? '' : 'off'}">${s.is_active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openStaffModal('${s.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeStaff('${s.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="7" style="color:var(--muted);text-align:center;padding:20px;">Belum ada data</td></tr>';
    }
    function openStaffModal(id) {
      const s = id ? state.staffNeeds.find(x => x.id === id) : { posisi: '', kualifikasi: '', jumlah: 1, batas_lamar: '', keterangan: '', is_active: true, form_file: null, form_file_name: null };
      pendingStaffFormFile = undefined; pendingStaffFormFileName = undefined;
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay"><div class="modal">
    <h3>${id ? 'Ubah Kebutuhan Tenaga' : 'Tambah Kebutuhan Tenaga'}</h3>
    <div class="field"><label>Posisi</label><input id="m-st-posisi" value="${escapeHtml(s.posisi)}" placeholder="mis. Bidan Desa, Tenaga Administrasi"></div>
    <div class="field"><label>Kualifikasi</label><input id="m-st-kualifikasi" value="${escapeHtml(s.kualifikasi)}"></div>
    <div class="field-row">
      <div class="field"><label>Jumlah</label><input id="m-st-jumlah" type="number" value="${s.jumlah}"></div>
      <div class="field"><label>Batas Lamar</label><input id="m-st-batas" type="date" value="${s.batas_lamar || ''}"></div>
    </div>
    <div class="field"><label>Keterangan</label><input id="m-st-ket" value="${escapeHtml(s.keterangan || '')}"></div>
    <div class="field"><label>Status</label>
      <select id="m-st-active"><option value="true" ${s.is_active ? 'selected' : ''}>Aktif</option><option value="false" ${!s.is_active ? 'selected' : ''}>Nonaktif</option></select>
    </div>
    <div class="field">
      <label>Formulir Pendaftaran (PDF/DOC, maks 5MB)</label>
      <div id="m-st-form-current" style="font-size:12px;color:var(--muted);margin-bottom:6px;">
        ${isSafeFileSrc(s.form_file) ? `📄 ${escapeHtml(s.form_file_name || 'formulir')} <a href="#" onclick="event.preventDefault();document.getElementById('m-st-form-current').style.display='none';pendingStaffFormFile=null;pendingStaffFormFileName=null;">(hapus)</a>` : 'Belum ada formulir diunggah'}
      </div>
      <input type="file" id="m-st-form-file" accept=".pdf,.doc,.docx">
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" id="m-st-save-btn" onclick="saveStaffModal('${id || ''}')">Simpan</button>
    </div>
  </div></div>`;
      document.getElementById('m-st-form-file').addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        try {
          pendingStaffFormFile = await fileToDataUrl(file, 5, ['.pdf', '.doc', '.docx']);
          pendingStaffFormFileName = file.name;
          toast('Formulir siap disimpan');
        } catch (err) { toast(errMsg(err), true); e.target.value = ''; }
      });
    }
    async function saveStaffModal(id) {
      const row = {
        id: id || undefined,
        posisi: document.getElementById('m-st-posisi').value.trim(),
        kualifikasi: document.getElementById('m-st-kualifikasi').value.trim(),
        jumlah: parseInt(document.getElementById('m-st-jumlah').value) || 1,
        batas_lamar: document.getElementById('m-st-batas').value || null,
        keterangan: document.getElementById('m-st-ket').value.trim(),
        is_active: document.getElementById('m-st-active').value === 'true'
      };
      if (!row.posisi) { toast('Posisi wajib diisi', true); return; }
      if (pendingStaffFormFile !== undefined) { row.form_file = pendingStaffFormFile; row.form_file_name = pendingStaffFormFileName; }
      const btn = document.getElementById('m-st-save-btn');
      btn.disabled = true;
      try { await DataLayer.upsertStaffNeed(row); closeModal(); await loadAdminData(); toast('Data tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); btn.disabled = false; }
    }
    async function removeStaff(id) {
      if (!confirm('Hapus data ini?')) return;
      try { await DataLayer.deleteStaffNeed(id); await loadAdminData(); toast('Data dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-staff').addEventListener('click', () => openStaffModal(null));

    /* ---- Kode Verifikasi Google Authenticator (TOTP) ---- */
    let totpPreviewTimer;
    function renderTotpPanel() {
      const el = document.getElementById('totp-panel');
      if (!el) return;
      const secret = state.settings.totp_secret || '';
      clearInterval(totpPreviewTimer);
      if (!secret) {
        el.innerHTML = `
    <p style="font-size:12.5px;color:var(--muted);margin:0 0 12px;">Belum ada kode verifikasi aktif. Buat kunci rahasia baru, lalu tambahkan secara manual di aplikasi Google Authenticator (pilih "Masukkan kunci penyiapan", tipe: Berbasis waktu).</p>
    <button class="btn btn-primary btn-sm" onclick="generateTotpSecret()">🔑 Buat Kode Verifikasi</button>`;
        return;
      }
      el.innerHTML = `
    <div class="vm-meta-row" style="max-width:420px;">
      <span class="k">Kunci Rahasia (Setup Key)</span>
      <span class="v" style="font-family:monospace;letter-spacing:1px;">${escapeHtml(secret)}</span>
    </div>
    <p style="font-size:12px;color:var(--muted);margin:10px 0;">Buka Google Authenticator → tambah akun → "Masukkan kunci penyiapan" → nama: <b>Puskesmas Puger</b>, kunci di atas, tipe <b>Berbasis waktu</b>, 6 digit.</p>
    <div class="vm-meta-row" style="max-width:220px;">
      <span class="k">Kode saat ini</span>
      <span class="v" id="totp-live-code" style="font-family:monospace;font-size:15px;letter-spacing:2px;">------</span>
    </div>
    <div class="modal-actions" style="justify-content:flex-start;margin-top:14px;">
      <button class="btn btn-danger btn-sm" onclick="generateTotpSecret()">🔄 Buat Ulang Kode Verifikasi</button>
    </div>`;
      const refresh = async () => { const code = await totpGenerate(secret); const c = document.getElementById('totp-live-code'); if (c) c.textContent = code; };
      refresh();
      totpPreviewTimer = setInterval(refresh, 1000);
    }
    async function generateTotpSecret() {
      if (state.settings.totp_secret && !confirm('Membuat ulang kode verifikasi akan membatalkan kunci lama di Google Authenticator. Lanjutkan?')) return;
      try {
        const secret = randomBase32Secret(16);
        await DataLayer.saveSetting('totp_secret', secret);
        state.settings.totp_secret = secret;
        renderTotpPanel();
        toast('Kode verifikasi baru dibuat — daftarkan ke Google Authenticator');
      } catch (e) { toast('Gagal membuat kode: ' + errMsg(e), true); }
    }

    /* ---- Berkas Pendaftaran Masuk ---- */
    function renderApplicationsTable() {
      const el = document.getElementById('applications-tbody');
      if (!el) return;
      el.innerHTML = (state.applications || []).map(a => `
    <tr>
      <td style="font-size:12px;">${a.created_at ? new Date(a.created_at).toLocaleString('id-ID') : '-'}</td>
      <td>${escapeHtml(a.posisi)}</td>
      <td><b>${escapeHtml(a.nama)}</b></td>
      <td>${escapeHtml(a.kontak)}</td>
      <td>${isSafeFileSrc(a.file_data) ? `<a class="btn btn-outline btn-sm" href="${a.file_data}" download="${escapeHtml(a.file_name || 'berkas')}">⬇ Unduh</a>` : '-'}</td>
      <td><button class="btn btn-danger btn-sm" onclick="removeApplication('${a.id}')">Hapus</button></td>
    </tr>`).join('') || '<tr><td colspan="6" style="color:var(--muted);text-align:center;padding:20px;">Belum ada berkas pendaftaran masuk</td></tr>';
    }
    async function removeApplication(id) {
      if (!confirm('Hapus berkas pendaftaran ini?')) return;
      try { await DataLayer.deleteApplication(id); await loadAdminData(); toast('Berkas dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }

    /* ---- SDM: Jadwal Kunjungan Dokter Spesialis ---- */
    function renderDoctorTable() {
      document.getElementById('doctor-tbody').innerHTML = state.doctorSchedule.map(s => `
    <tr>
      <td><b>${escapeHtml(s.nama_dokter)}</b></td>
      <td>${escapeHtml(s.spesialisasi)}</td>
      <td>${escapeHtml(s.hari)}</td>
      <td>${escapeHtml(s.jam)}</td>
      <td>${escapeHtml(s.lokasi || '-')}</td>
      <td><span class="badge ${s.is_active ? '' : 'off'}">${s.is_active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openDoctorModal('${s.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeDoctor('${s.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="7" style="color:var(--muted);text-align:center;padding:20px;">Belum ada data</td></tr>';
    }
    function openDoctorModal(id) {
      const s = id ? state.doctorSchedule.find(x => x.id === id) : { nama_dokter: '', spesialisasi: '', hari: '', jam: '', lokasi: '', keterangan: '', is_active: true };
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay"><div class="modal">
    <h3>${id ? 'Ubah Jadwal Dokter' : 'Tambah Jadwal Dokter'}</h3>
    <div class="field"><label>Nama Dokter</label><input id="m-d-nama" value="${escapeHtml(s.nama_dokter)}" placeholder="mis. dr. Ahmad Fauzi, Sp.PD"></div>
    <div class="field"><label>Spesialisasi</label><input id="m-d-spesialisasi" value="${escapeHtml(s.spesialisasi)}"></div>
    <div class="field-row">
      <div class="field"><label>Hari Praktik</label><input id="m-d-hari" value="${escapeHtml(s.hari)}" placeholder="mis. Selasa & Kamis"></div>
      <div class="field"><label>Jam</label><input id="m-d-jam" value="${escapeHtml(s.jam)}" placeholder="mis. 09.00–12.00"></div>
    </div>
    <div class="field"><label>Lokasi / Poli</label><input id="m-d-lokasi" value="${escapeHtml(s.lokasi || '')}"></div>
    <div class="field"><label>Keterangan</label><input id="m-d-ket" value="${escapeHtml(s.keterangan || '')}"></div>
    <div class="field"><label>Status</label>
      <select id="m-d-active"><option value="true" ${s.is_active ? 'selected' : ''}>Aktif</option><option value="false" ${!s.is_active ? 'selected' : ''}>Nonaktif</option></select>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="saveDoctorModal('${id || ''}')">Simpan</button>
    </div>
  </div></div>`;
    }
    async function saveDoctorModal(id) {
      const row = {
        id: id || undefined,
        nama_dokter: document.getElementById('m-d-nama').value.trim(),
        spesialisasi: document.getElementById('m-d-spesialisasi').value.trim(),
        hari: document.getElementById('m-d-hari').value.trim(),
        jam: document.getElementById('m-d-jam').value.trim(),
        lokasi: document.getElementById('m-d-lokasi').value.trim(),
        keterangan: document.getElementById('m-d-ket').value.trim(),
        is_active: document.getElementById('m-d-active').value === 'true'
      };
      if (!row.nama_dokter) { toast('Nama dokter wajib diisi', true); return; }
      try { await DataLayer.upsertDoctorSchedule(row); closeModal(); await loadAdminData(); toast('Jadwal dokter tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeDoctor(id) {
      if (!confirm('Hapus jadwal ini?')) return;
      try { await DataLayer.deleteDoctorSchedule(id); await loadAdminData(); toast('Jadwal dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-doctor').addEventListener('click', () => openDoctorModal(null));

    /* ---- Hotline & Kontak ---- */
    function renderHotlineTable() {
      document.getElementById('hotline-tbody').innerHTML = state.hotline.map(h => `
    <tr>
      <td><b>${escapeHtml(h.nama)}</b>${h.keterangan ? `<div style="font-size:11.5px;color:var(--muted);margin-top:2px;">${escapeHtml(h.keterangan)}</div>` : ''}</td>
      <td>${getHotlineTarget(h.nomor).icon} ${escapeHtml(h.nomor)}</td>
      <td><span class="badge ${h.is_active ? '' : 'off'}">${h.is_active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openHotlineModal('${h.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeHotline('${h.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:20px;">Belum ada data</td></tr>';
    }
    function openHotlineModal(id) {
      const h = id ? state.hotline.find(x => x.id === id) : { nama: '', nomor: '', keterangan: '', is_active: true };
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay"><div class="modal">
    <h3>${id ? 'Ubah Hotline' : 'Tambah Hotline'}</h3>
    <div class="field"><label>Layanan Call Inbound</label><input id="m-h-nama" value="${escapeHtml(h.nama)}" placeholder="mis. Hotline Homecare"></div>
    <div class="field">
      <label>Nomor / Link Hotline</label>
      <input id="m-h-nomor" value="${escapeHtml(h.nomor)}" placeholder="mis. 0812-3456-7890, wa.me/62812..., https://... atau email">
      <div style="font-size:11px;color:var(--muted);margin-top:4px;">Bisa diisi nomor telepon (📞), link WhatsApp (💬), link web/Instagram (🔗), atau email (✉️) — sistem otomatis mengarahkan sesuai jenisnya.</div>
    </div>
    <div class="field"><label>Keterangan</label><input id="m-h-ket" value="${escapeHtml(h.keterangan || '')}" placeholder="opsional"></div>
    <div class="field"><label>Status</label>
      <select id="m-h-active"><option value="true" ${h.is_active ? 'selected' : ''}>Aktif</option><option value="false" ${!h.is_active ? 'selected' : ''}>Nonaktif</option></select>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="saveHotlineModal('${id || ''}')">Simpan</button>
    </div>
  </div></div>`;
    }
    async function saveHotlineModal(id) {
      const row = {
        id: id || undefined,
        nama: document.getElementById('m-h-nama').value.trim(),
        nomor: document.getElementById('m-h-nomor').value.trim(),
        keterangan: document.getElementById('m-h-ket').value.trim(),
        is_active: document.getElementById('m-h-active').value === 'true'
      };
      if (!row.nama) { toast('Nama layanan wajib diisi', true); return; }
      if (!row.nomor) { toast('Nomor hotline wajib diisi', true); return; }
      try { await DataLayer.upsertHotline(row); closeModal(); await loadAdminData(); toast('Data hotline tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeHotline(id) {
      if (!confirm('Hapus data hotline ini?')) return;
      try { await DataLayer.deleteHotline(id); await loadAdminData(); toast('Data dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-hotline').addEventListener('click', () => openHotlineModal(null));

    /* ---- Breaking News (Poster / Flyer) ---- */
    /* Menyimpan sementara data URL poster yang baru dipilih di modal tambah/ubah,
       sebelum tombol "Simpan" ditekan — mengikuti pola pendingLogoDataUrl di atas.
       undefined = belum diubah (pertahankan gambar lama saat mengubah data). */
    let pendingBreakingImageDataUrl;
    function renderBreakingNewsTable() {
      document.getElementById('breaking-tbody').innerHTML = state.breakingNews.map(n => `
    <tr>
      <td>${isSafeImageSrc(n.image) ? `<img class="poster-thumb" src="${escapeHtml(n.image)}" alt="">` : '<span style="color:var(--muted);font-size:11px;">Belum ada gambar</span>'}</td>
      <td><b>${escapeHtml(n.title || '(Tanpa judul)')}</b></td>
      <td><span class="badge ${n.is_active ? '' : 'off'}">${n.is_active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openBreakingModal('${n.id}')">Ubah</button>
        <button class="btn btn-danger btn-sm" onclick="removeBreakingNews('${n.id}')">Hapus</button>
      </div></td>
    </tr>`).join('') || '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:20px;">Belum ada poster/flyer</td></tr>';
    }
    function openBreakingModal(id) {
      const n = id ? state.breakingNews.find(x => x.id === id) : { title: '', image: '', is_active: true };
      pendingBreakingImageDataUrl = undefined;
      document.getElementById('modal-root').innerHTML = `
  <div class="modal-overlay"><div class="modal">
    <h3>${id ? 'Ubah Poster/Flyer' : 'Tambah Poster/Flyer'}</h3>
    <div class="field"><label>Gambar Poster/Flyer</label>
      <div style="display:flex;align-items:center;gap:14px;">
        <div class="poster-preview-wrap"><img id="m-b-preview" src="${isSafeImageSrc(n.image) ? escapeHtml(n.image) : ''}" alt="Pratinjau poster"></div>
        <div>
          <input type="file" id="m-b-file" accept="image/png,image/jpeg,image/webp" class="hidden">
          <button type="button" class="btn btn-outline btn-sm" id="m-b-pick">Pilih Gambar</button>
          <p class="desc" style="margin:6px 0 0;">Format JPG/PNG/WebP, maksimal 8MB.</p>
        </div>
      </div>
    </div>
    <div class="field"><label>Judul (opsional)</label><input id="m-b-title" value="${escapeHtml(n.title || '')}" placeholder="mis. Jadwal Vaksinasi Booster"></div>
    <div class="field"><label>Status</label>
      <select id="m-b-active"><option value="true" ${n.is_active ? 'selected' : ''}>Aktif</option><option value="false" ${!n.is_active ? 'selected' : ''}>Nonaktif</option></select>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="saveBreakingModal('${id || ''}')">Simpan</button>
    </div>
  </div></div>`;
      document.getElementById('m-b-pick').addEventListener('click', () => document.getElementById('m-b-file').click());
      document.getElementById('m-b-file').addEventListener('change', async (e) => {
        const file = e.target.files[0]; if (!file) return;
        try {
          const dataUrl = await fileToResizedDataUrl(file, 1400);
          pendingBreakingImageDataUrl = dataUrl;
          document.getElementById('m-b-preview').src = dataUrl;
        } catch (err) { toast('Gagal memproses gambar: ' + errMsg(err), true); }
      });
    }
    async function saveBreakingModal(id) {
      const existing = id ? state.breakingNews.find(x => x.id === id) : null;
      const image = pendingBreakingImageDataUrl !== undefined ? pendingBreakingImageDataUrl : ((existing && existing.image) || '');
      const row = {
        id: id || undefined,
        title: document.getElementById('m-b-title').value.trim(),
        image: image,
        is_active: document.getElementById('m-b-active').value === 'true'
      };
      if (!row.image) { toast('Gambar poster/flyer wajib diunggah', true); return; }
      try { await DataLayer.upsertBreakingNews(row); closeModal(); await loadAdminData(); toast('Poster/flyer tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }
    async function removeBreakingNews(id) {
      if (!confirm('Hapus poster/flyer ini?')) return;
      try { await DataLayer.deleteBreakingNews(id); await loadAdminData(); toast('Poster/flyer dihapus'); }
      catch (e) { toast('Gagal menghapus: ' + errMsg(e), true); }
    }
    document.getElementById('btn-add-breaking').addEventListener('click', () => openBreakingModal(null));

    /* ---- Jam Layanan ---- */
    function renderHoursTable() {
      document.getElementById('hours-tbody').innerHTML = state.hours.map(h => `
    <tr data-day="${h.day_order}">
      <td><b>${h.day_name}</b></td>
      <td><label style="display:flex;align-items:center;gap:6px;font-size:12.5px;">
        <input type="checkbox" class="h-closed" ${h.is_closed ? 'checked' : ''}> Libur</label></td>
      <td><input type="time" class="h-open" value="${h.open_time || ''}" style="border:1px solid var(--border);border-radius:8px;padding:6px 8px;" ${h.is_closed ? 'disabled' : ''}></td>
      <td><input type="time" class="h-close" value="${h.close_time || ''}" style="border:1px solid var(--border);border-radius:8px;padding:6px 8px;" ${h.is_closed ? 'disabled' : ''}></td>
      <td><input type="text" class="h-note" value="${escapeHtml(h.note || '')}" placeholder="Catatan (opsional)" style="border:1px solid var(--border);border-radius:8px;padding:6px 8px;width:150px;"></td>
      <td><button class="btn btn-primary btn-sm" onclick="saveHourRow(${h.day_order})">Simpan</button></td>
    </tr>`).join('');

      document.querySelectorAll('.h-closed').forEach(cb => {
        cb.addEventListener('change', (e) => {
          const tr = e.target.closest('tr');
          tr.querySelectorAll('.h-open,.h-close').forEach(inp => inp.disabled = e.target.checked);
        });
      });
    }
    async function saveHourRow(dayOrder) {
      const tr = document.querySelector(`#hours-tbody tr[data-day="${dayOrder}"]`);
      const is_closed = tr.querySelector('.h-closed').checked;
      const row = {
        day_order: dayOrder,
        is_closed,
        open_time: is_closed ? null : (tr.querySelector('.h-open').value || null),
        close_time: is_closed ? null : (tr.querySelector('.h-close').value || null),
        note: tr.querySelector('.h-note').value.trim() || null
      };
      try { await DataLayer.saveHour(row); await loadAdminData(); toast('Jam layanan tersimpan'); }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    }

    /* ---- Pengaturan Tampilan ---- */
    function renderTampilanForm() {
      const profile = state.settings.profile || {};
      pendingLogoDataUrl = undefined;
      document.getElementById('set-logo-preview').src = isSafeImageSrc(profile.logo) ? profile.logo : DEFAULT_PROFILE_LOGO;
      document.getElementById('set-name').value = profile.name || '';
      document.getElementById('set-subtitle').value = profile.subtitle || '';
      document.getElementById('set-tagline').value = profile.tagline || '';
      document.getElementById('set-address').value = profile.address || '';
      document.getElementById('set-phone').value = profile.phone || '';
      document.getElementById('set-emergency').value = profile.emergency || '';
      document.getElementById('set-maps').value = profile.mapsUrl || '';

      const theme = state.settings.theme || {};
      document.getElementById('col-primary').value = theme.primary || '#EC9CB6';
      document.getElementById('col-primary-dark').value = theme.primaryDark || '#C26A88';
      document.getElementById('col-accent').value = theme.accent || '#F4B8CE';
      document.getElementById('col-bg').value = theme.bg || '#FFF6F9';
      document.getElementById('ticker-speed').value = theme.tickerSpeed || 32;

      /* ---- Panel warna grafik capaian ---- */
      const chart = { ...defaultChartSettings(), ...(theme.chart || {}) };
      document.getElementById('cc-threshold-excellent').value = chart.thresholdExcellent;
      document.getElementById('cc-threshold-good').value = chart.thresholdGood;
      document.getElementById('cc-color-excellent').value = chart.colorExcellent;
      document.getElementById('cc-color-good').value = chart.colorGood;
      document.getElementById('cc-color-poor').value = chart.colorPoor;
      document.getElementById('cc-color-fixed').value = chart.colorFixed;
      let chartMode = chart.mode || 'auto';
      const setChartModeUI = () => {
        document.querySelectorAll('[data-chartmode]').forEach(b => {
          b.classList.toggle('btn-primary', b.dataset.chartmode === chartMode);
          b.classList.toggle('btn-outline', b.dataset.chartmode !== chartMode);
        });
        document.getElementById('chart-color-auto-fields').classList.toggle('hidden', chartMode !== 'auto');
        document.getElementById('chart-color-fixed-fields').classList.toggle('hidden', chartMode !== 'fixed');
      };
      setChartModeUI();
      document.querySelectorAll('[data-chartmode]').forEach(b => {
        b.onclick = () => { chartMode = b.dataset.chartmode; setChartModeUI(); };
      });
      document.getElementById('btn-save-chartcolor').onclick = async () => {
        const newChart = {
          mode: chartMode,
          thresholdExcellent: parseFloat(document.getElementById('cc-threshold-excellent').value) || 100,
          thresholdGood: parseFloat(document.getElementById('cc-threshold-good').value) || 75,
          colorExcellent: document.getElementById('cc-color-excellent').value,
          colorGood: document.getElementById('cc-color-good').value,
          colorPoor: document.getElementById('cc-color-poor').value,
          colorFixed: document.getElementById('cc-color-fixed').value
        };
        try {
          const updatedTheme = { ...(state.settings.theme || {}), chart: newChart };
          await DataLayer.saveSetting('theme', updatedTheme);
          state.settings.theme = updatedTheme;
          toast('Warna grafik capaian tersimpan');
        } catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
      };

      document.getElementById('preset-grid').innerHTML = THEME_PRESETS.map(p => `
    <div class="theme-preset" data-preset="${p.name}">
      <span class="dots"><span style="background:${p.primary}"></span><span style="background:${p.accent}"></span><span style="background:${p.primaryDark}"></span></span>
      ${p.name}
    </div>`).join('');
      document.querySelectorAll('.theme-preset').forEach(el => {
        el.addEventListener('click', () => {
          const preset = THEME_PRESETS.find(p => p.name === el.dataset.preset);
          document.getElementById('col-primary').value = preset.primary;
          document.getElementById('col-primary-dark').value = preset.primaryDark;
          document.getElementById('col-accent').value = preset.accent;
          document.getElementById('col-bg').value = preset.bg;
          applyTheme({ ...preset, layout: currentLayoutChoice });
          document.querySelectorAll('.theme-preset').forEach(x => x.classList.remove('selected'));
          el.classList.add('selected');
        });
      });

      let currentLayoutChoice = theme.layout || 'comfortable';
      document.querySelectorAll('.layout-toggle button').forEach(b => {
        b.classList.toggle('btn-primary', b.dataset.layout === currentLayoutChoice);
        b.classList.toggle('btn-outline', b.dataset.layout !== currentLayoutChoice);
        b.onclick = () => {
          currentLayoutChoice = b.dataset.layout;
          document.querySelectorAll('.layout-toggle button').forEach(x => { x.classList.remove('btn-primary'); x.classList.add('btn-outline'); });
          b.classList.remove('btn-outline'); b.classList.add('btn-primary');
          applyTheme({ ...readColorInputs(), layout: currentLayoutChoice });
        };
      });

      ['col-primary', 'col-primary-dark', 'col-accent', 'col-bg'].forEach(id => {
        document.getElementById(id).addEventListener('input', () => applyTheme({ ...readColorInputs(), layout: currentLayoutChoice }));
      });

      document.getElementById('ticker-speed').addEventListener('input', () => applyTheme({ ...readColorInputs(), layout: currentLayoutChoice }));

      document.getElementById('btn-save-theme').onclick = async () => {
        const theme = { ...(state.settings.theme || {}), ...readColorInputs(), layout: currentLayoutChoice, font: 'jakarta' };
        try { await DataLayer.saveSetting('theme', theme); state.settings.theme = theme; toast('Tampilan tersimpan'); }
        catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
      };
      document.getElementById('btn-reset-theme').onclick = () => {
        const def = THEME_PRESETS[0];
        document.getElementById('col-primary').value = def.primary;
        document.getElementById('col-primary-dark').value = def.primaryDark;
        document.getElementById('col-accent').value = def.accent;
        document.getElementById('col-bg').value = def.bg;
        document.getElementById('ticker-speed').value = 32;
        currentLayoutChoice = 'comfortable';
        applyTheme({ ...def, layout: 'comfortable', tickerSpeed: 32 });
      };
    }
    function readColorInputs() {
      return {
        primary: document.getElementById('col-primary').value,
        primaryDark: document.getElementById('col-primary-dark').value,
        accent: document.getElementById('col-accent').value,
        bg: document.getElementById('col-bg').value,
        tickerSpeed: parseInt(document.getElementById('ticker-speed').value) || 32
      };
    }
    document.getElementById('btn-pick-logo').addEventListener('click', () => document.getElementById('set-logo-file').click());
    document.getElementById('set-logo-file').addEventListener('change', async (e) => {
      const file = e.target.files[0];
      e.target.value = '';
      if (!file) return;
      try {
        const dataUrl = await fileToResizedDataUrl(file, 320);
        pendingLogoDataUrl = dataUrl;
        document.getElementById('set-logo-preview').src = dataUrl;
      } catch (err) { toast('Gagal memproses gambar: ' + errMsg(err), true); }
    });
    document.getElementById('btn-remove-logo').addEventListener('click', () => {
      pendingLogoDataUrl = null;
      document.getElementById('set-logo-preview').src = DEFAULT_PROFILE_LOGO;
    });
    document.getElementById('btn-save-profile').addEventListener('click', async () => {
      const currentLogo = (state.settings.profile && state.settings.profile.logo) || '';
      const logo = pendingLogoDataUrl === undefined ? currentLogo : (pendingLogoDataUrl || '');
      const profile = {
        name: document.getElementById('set-name').value.trim(),
        subtitle: document.getElementById('set-subtitle').value.trim(),
        tagline: document.getElementById('set-tagline').value.trim(),
        address: document.getElementById('set-address').value.trim(),
        phone: document.getElementById('set-phone').value.trim(),
        emergency: document.getElementById('set-emergency').value.trim(),
        mapsUrl: document.getElementById('set-maps').value.trim(),
        logo
      };
      try {
        await DataLayer.saveSetting('profile', profile);
        state.settings.profile = profile;
        pendingLogoDataUrl = undefined;
        applyProfileLogo(profile.logo);
        toast('Profil tersimpan');
      }
      catch (e) { toast('Gagal menyimpan: ' + errMsg(e), true); }
    });

    /* ---- Akun Admin ---- */
    function renderAkunForm() {
      const noteEl = document.getElementById('akun-demo-note');
      noteEl.classList.toggle('hidden', !DEMO_MODE);
      document.getElementById('akun-current-email').value =
        (state.session && state.session.user && state.session.user.email) || '';
      document.getElementById('akun-new-email').value = '';
      document.getElementById('akun-new-password').value = '';
      document.getElementById('akun-confirm-password').value = '';
      document.getElementById('akun-msg').innerHTML = '';

      document.getElementById('tambah-akun-demo-note').classList.toggle('hidden', !DEMO_MODE);
      document.getElementById('ta-email').value = '';
      document.getElementById('ta-password').value = '';
      document.getElementById('ta-confirm-password').value = '';
      document.getElementById('tambah-akun-msg').innerHTML = '';

      /* Hanya akun berperan Admin yang boleh mengelola akun lain / menambah akun baru.
         Operator hanya boleh mengubah email & password akunnya sendiri (panel di atas). */
      const myRole = (state.myProfile && state.myProfile.role) || 'admin';
      const isAdmin = DEMO_MODE || myRole === 'admin';
      document.getElementById('tambah-akun-panel').classList.toggle('hidden', !isAdmin);
    }

    /* ---- Kelola Pengguna & Peran (role + nonaktifkan akun) ---- */
    /* Menyimpan hasil fetch terakhir supaya pencarian di kotak "Cari berdasarkan
       email..." tidak perlu fetch ulang ke Supabase setiap kali mengetik —
       cukup filter ulang array yang sudah ada lalu render ulang tabelnya. */
    let lastPenggunaProfiles = [];

    async function renderPenggunaTable() {
      const note = document.getElementById('pengguna-demo-note');
      const msgEl = document.getElementById('pengguna-msg');
      const tbody = document.getElementById('pengguna-tbody');
      note.classList.toggle('hidden', !DEMO_MODE);
      msgEl.innerHTML = '';

      if (DEMO_MODE) {
        tbody.innerHTML = '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:20px;">Fitur ini aktif setelah aplikasi terhubung ke Supabase.</td></tr>';
        return;
      }

      let profiles;
      try {
        profiles = await DataLayer.getProfiles();
      } catch (err) {
        tbody.innerHTML = '';
        msgEl.innerHTML = `<div class="form-msg err">Gagal memuat daftar pengguna: ${escapeHtml(errMsg(err))}. Pastikan tabel <code>admin_profiles</code> sudah dibuat (lihat komentar CONFIG di awal berkas).</div>`;
        return;
      }

      const myRole = (state.myProfile && state.myProfile.role) || 'admin';
      const isAdmin = myRole === 'admin';
      if (!isAdmin) {
        msgEl.innerHTML = `<div class="form-msg" style="background:#EAF1FB;color:#2F5FA8;">Anda masuk sebagai <b>Operator</b>. Hanya akun berperan <b>Admin</b> yang dapat mengubah peran atau menonaktifkan akun.</div>`;
      }

      lastPenggunaProfiles = profiles;
      renderPenggunaRows(profiles);
    }

    /* Merender baris tabel dari array profiles yang sudah ada di memori (tanpa fetch
       ulang) — dipakai baik oleh renderPenggunaTable() (data baru) maupun oleh
       kotak pencarian (data lama, difilter ulang). */
    function renderPenggunaRows(profiles) {
      const tbody = document.getElementById('pengguna-tbody');
      const myId = state.session && state.session.user && state.session.user.id;
      const myRole = (state.myProfile && state.myProfile.role) || 'admin';
      const isAdmin = myRole === 'admin';

      if (!profiles.length) {
        tbody.innerHTML = '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:20px;">Belum ada data pengguna.</td></tr>';
        return;
      }

      const q = (document.getElementById('pengguna-search').value || '').trim().toLowerCase();
      const filtered = q ? profiles.filter(p => (p.email || '').toLowerCase().includes(q)) : profiles;
      if (!filtered.length) {
        tbody.innerHTML = '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:20px;">Tidak ada akun yang cocok dengan pencarian.</td></tr>';
        return;
      }

      /* Jaga-jaga agar tidak ada yang menonaktifkan/menurunkan peran satu-satunya Admin aktif
         yang tersisa — kalau itu terjadi, tidak akan ada lagi akun yang bisa mengelola akun lain.
         Dihitung dari SELURUH profiles (bukan hasil filter), supaya proteksi ini tidak berubah
         hanya karena sedang mencari sesuatu di kotak pencarian. */
      const activeAdminCount = profiles.filter(p => p.role === 'admin' && p.is_active !== false).length;

      tbody.innerHTML = filtered.map(p => {
        const isMe = p.id === myId;
        const active = p.is_active !== false;
        const isSoleActiveAdmin = active && p.role === 'admin' && activeAdminCount <= 1;
        const roleLocked = !isAdmin || isSoleActiveAdmin;
        const toggleLocked = !isAdmin || isMe;
        const emailEsc = escapeHtml((p.email || '').replace(/'/g, "\\'"));
        return `
    <tr data-id="${escapeHtml(p.id)}">
      <td>${escapeHtml(p.email || '(tanpa email)')}${isMe ? ' <span style="color:var(--muted);font-size:11px;">(Anda)</span>' : ''}</td>
      <td>
        <select class="pengguna-role-select" ${roleLocked ? 'disabled' : ''} onchange="changePenggunaRole('${p.id}', this.value)" title="${isSoleActiveAdmin ? 'Tidak dapat mengubah peran satu-satunya Admin aktif' : ''}">
          <option value="admin" ${p.role === 'admin' ? 'selected' : ''}>Admin</option>
          <option value="operator" ${p.role !== 'admin' ? 'selected' : ''}>Operator</option>
        </select>
      </td>
      <td><span class="badge ${active ? '' : 'off'}">${active ? 'Aktif' : 'Nonaktif'}</span></td>
      <td><div class="row-actions">
        <button class="btn ${active ? 'btn-danger' : 'btn-outline'} btn-sm" ${toggleLocked ? 'disabled' : ''} onclick="togglePenggunaActive('${p.id}', ${active}, '${emailEsc}')" title="${isMe ? 'Tidak dapat menonaktifkan akun sendiri' : ''}">${active ? 'Nonaktifkan' : 'Aktifkan'}</button>
        <button class="btn btn-outline btn-sm" ${!isAdmin ? 'disabled' : ''} onclick="resetPenggunaPassword('${p.id}', '${emailEsc}')" title="Kirim tautan reset password ke email akun ini">Reset Password</button>
      </div></td>
    </tr>`;
      }).join('');
    }
    document.getElementById('pengguna-search').addEventListener('input', () => renderPenggunaRows(lastPenggunaProfiles));

    async function changePenggunaRole(id, newRole) {
      const select = document.querySelector(`#pengguna-tbody tr[data-id="${CSS.escape(id)}"] select.pengguna-role-select`);
      if (select) select.disabled = true;
      const target = lastPenggunaProfiles.find(p => p.id === id);
      const oldRole = target ? target.role : '';
      /* BUGFIX: jika admin yang sedang login menurunkan perannya sendiri menjadi Operator,
         panel "Tambah Akun Admin Baru" (yang visibilitasnya diatur di renderAkunForm, bukan
         renderPenggunaTable) sebelumnya tetap tampil sampai halaman dimuat ulang — meski
         akun tsb seharusnya sudah kehilangan akses tersebut seketika. */
      const isSelf = id === (state.session && state.session.user && state.session.user.id);
      try {
        await DataLayer.updateProfileRole(id, newRole);
        if (isSelf) {
          state.myProfile = { ...(state.myProfile || {}), role: newRole };
        }
        await DataLayer.logAudit('change_role', (target && target.email) || '', `${oldRole || '?'} → ${newRole}`);
        toast('Peran pengguna diperbarui');
      } catch (err) {
        document.getElementById('pengguna-msg').innerHTML = `<div class="form-msg err">Gagal mengubah peran: ${escapeHtml(errMsg(err))}</div>`;
      } finally {
        await renderPenggunaTable();
        if (isSelf) renderAkunForm();
        renderAuditLog();
      }
    }

    async function togglePenggunaActive(id, currentlyActive, email) {
      const nextActive = !currentlyActive;
      if (!nextActive && !confirm(`Nonaktifkan akun "${email}"? Akun ini tidak akan bisa masuk ke panel admin sampai diaktifkan kembali.`)) {
        return;
      }
      const row = document.querySelector(`#pengguna-tbody tr[data-id="${CSS.escape(id)}"]`);
      if (row) row.querySelectorAll('button, select').forEach(el => el.disabled = true);
      try {
        await DataLayer.updateProfileActive(id, nextActive);
        await DataLayer.logAudit(nextActive ? 'activate_account' : 'deactivate_account', email, '');
        toast(nextActive ? 'Akun diaktifkan kembali' : 'Akun dinonaktifkan');
      } catch (err) {
        document.getElementById('pengguna-msg').innerHTML = `<div class="form-msg err">Gagal mengubah status akun: ${escapeHtml(errMsg(err))}</div>`;
      } finally {
        await renderPenggunaTable();
        renderAuditLog();
      }
    }

    /* ---- Reset Password akun lain (oleh Admin) ----
       Tidak memakai service_role key (tidak aman disimpan di file ini), sehingga
       Admin TIDAK BISA langsung mengganti password akun lain begitu saja. Yang
       dilakukan di sini adalah mengirim tautan reset password resmi dari Supabase
       ke email akun tersebut — pemilik akun tetap yang menentukan password barunya
       lewat tautan itu. Ini cara yang aman & didukung penuh oleh Supabase Auth. */
    async function resetPenggunaPassword(id, email) {
      if (!email) { toast('Akun ini tidak punya alamat email', true); return; }
      if (!confirm(`Kirim tautan reset password ke "${email}"?`)) return;
      const btn = document.querySelector(`#pengguna-tbody tr[data-id="${CSS.escape(id)}"] .row-actions button:last-child`);
      if (btn) btn.disabled = true;
      try {
        const { error } = await sb.auth.resetPasswordForEmail(email);
        if (error) throw error;
        await DataLayer.logAudit('reset_password_request', email, '');
        toast(`Tautan reset password terkirim ke ${email}`);
      } catch (err) {
        document.getElementById('pengguna-msg').innerHTML = `<div class="form-msg err">Gagal mengirim tautan reset password: ${escapeHtml(errMsg(err))}</div>`;
      } finally {
        await renderPenggunaTable();
        renderAuditLog();
      }
    }

    /* ---- Riwayat Aktivitas Akun (audit log) ---- */
    const AUDIT_ACTION_LABELS = {
      change_role: 'Ubah peran',
      activate_account: 'Aktifkan akun',
      deactivate_account: 'Nonaktifkan akun',
      create_account: 'Tambah akun baru',
      reset_password_request: 'Minta reset password'
    };
    async function renderAuditLog() {
      const panel = document.getElementById('audit-log-panel');
      const note = document.getElementById('audit-log-demo-note');
      const tbody = document.getElementById('audit-log-tbody');
      const myRole = (state.myProfile && state.myProfile.role) || 'admin';
      const isAdmin = DEMO_MODE || myRole === 'admin';
      /* Riwayat ini hanya relevan/berguna untuk Admin (yang berhak mengelola akun lain);
         Operator tidak melihat panel ini sama sekali. */
      panel.classList.toggle('hidden', !isAdmin);
      if (!isAdmin) return;

      note.classList.toggle('hidden', !DEMO_MODE);
      if (DEMO_MODE) {
        tbody.innerHTML = '<tr><td colspan="5" style="color:var(--muted);text-align:center;padding:20px;">Fitur ini aktif setelah aplikasi terhubung ke Supabase.</td></tr>';
      }
      try {
        const rows = await DataLayer.getAuditLog(30);
        if (!rows.length) {
          tbody.innerHTML = '<tr><td colspan="5" style="color:var(--muted);text-align:center;padding:20px;">Belum ada riwayat aktivitas.</td></tr>';
          return;
        }
        tbody.innerHTML = rows.map(r => `
    <tr>
      <td style="white-space:nowrap;">${escapeHtml(r.created_at ? new Date(r.created_at).toLocaleString('id-ID') : '-')}</td>
      <td>${escapeHtml(r.actor_email || '-')}</td>
      <td>${escapeHtml(AUDIT_ACTION_LABELS[r.action] || r.action)}</td>
      <td>${escapeHtml(r.target_email || '-')}</td>
      <td>${escapeHtml(r.detail || '-')}</td>
    </tr>`).join('');
      } catch (err) {
        tbody.innerHTML = `<tr><td colspan="5" style="color:var(--muted);text-align:center;padding:20px;">Riwayat belum tersedia. Pastikan tabel <code>admin_audit_log</code> sudah dibuat (lihat komentar CONFIG di awal berkas).</td></tr>`;
      }
    }
    document.getElementById('btn-save-akun').addEventListener('click', async () => {
      const msg = document.getElementById('akun-msg');
      msg.innerHTML = '';
      const newEmail = document.getElementById('akun-new-email').value.trim();
      const newPassword = document.getElementById('akun-new-password').value;
      const confirmPassword = document.getElementById('akun-confirm-password').value;

      if (!newEmail && !newPassword) {
        msg.innerHTML = `<div class="form-msg err">Isi email baru dan/atau password baru terlebih dahulu.</div>`;
        return;
      }
      if (newPassword && newPassword.length < 6) {
        msg.innerHTML = `<div class="form-msg err">Password baru minimal 6 karakter.</div>`;
        return;
      }
      if (newPassword && newPassword !== confirmPassword) {
        msg.innerHTML = `<div class="form-msg err">Konfirmasi password baru tidak cocok.</div>`;
        return;
      }
      if (DEMO_MODE) {
        msg.innerHTML = `<div class="form-msg err">Fitur edit akun hanya aktif setelah aplikasi terhubung ke Supabase.</div>`;
        return;
      }
      try {
        const payload = {};
        if (newEmail) payload.email = newEmail;
        if (newPassword) payload.password = newPassword;
        const { error } = await sb.auth.updateUser(payload);
        if (error) throw error;
        const { data } = await sb.auth.getSession();
        state.session = data.session;
        document.getElementById('akun-current-email').value = (state.session && state.session.user && state.session.user.email) || newEmail;
        document.getElementById('akun-new-email').value = '';
        document.getElementById('akun-new-password').value = '';
        document.getElementById('akun-confirm-password').value = '';
        msg.innerHTML = `<div class="form-msg ok">Akun berhasil diperbarui${newEmail ? '. Jika email diubah, cek kotak masuk email baru untuk konfirmasi.' : '.'}</div>`;
      } catch (err) {
        msg.innerHTML = `<div class="form-msg err">Gagal memperbarui akun: ${escapeHtml(errMsg(err))}</div>`;
      }
    });

    document.getElementById('btn-tambah-akun').addEventListener('click', async () => {
      const msg = document.getElementById('tambah-akun-msg');
      msg.innerHTML = '';
      const email = document.getElementById('ta-email').value.trim();
      const password = document.getElementById('ta-password').value;
      const confirmPassword = document.getElementById('ta-confirm-password').value;
      /* BUGFIX: sebelumnya nilai peran yang dipilih di sini tidak pernah dipakai/disimpan,
         sehingga akun baru selalu berakhir tanpa baris admin_profiles (peran baru diketahui
         setelah akun tsb login sendiri lewat self-heal, dan selalu berperan default "admin"
         meski Operator yang dipilih). Sekarang perannya benar-benar disimpan ke admin_profiles. */
      const role = document.getElementById('ta-role').value === 'operator' ? 'operator' : 'admin';

      if (!email || !password) {
        msg.innerHTML = `<div class="form-msg err">Isi email dan password akun baru terlebih dahulu.</div>`;
        return;
      }
      if (password.length < 6) {
        msg.innerHTML = `<div class="form-msg err">Password minimal 6 karakter.</div>`;
        return;
      }
      if (password !== confirmPassword) {
        msg.innerHTML = `<div class="form-msg err">Konfirmasi password tidak cocok.</div>`;
        return;
      }
      if (DEMO_MODE) {
        msg.innerHTML = `<div class="form-msg err">Fitur ini hanya aktif setelah aplikasi terhubung ke Supabase.</div>`;
        return;
      }
      if (state.myProfile && state.myProfile.role !== 'admin') {
        msg.innerHTML = `<div class="form-msg err">Hanya akun berperan Admin yang dapat menambah akun baru.</div>`;
        return;
      }
      const btn = document.getElementById('btn-tambah-akun');
      btn.disabled = true;
      try {
        /* Pakai sbForNewAdmin (klien terpisah) supaya sesi admin yang sedang login tidak
           tergantikan oleh sesi akun baru ini. */
        const { data, error } = await sbForNewAdmin.auth.signUp({ email, password });
        if (error) throw error;
        /* BUGFIX: jika email sudah terdaftar sebelumnya, Supabase tetap mengembalikan objek
           "user" tanpa error (agar tidak bocorkan email mana yang sudah terdaftar), tapi
           array identities-nya kosong. Tanpa pengecekan ini, aplikasi salah melapor "berhasil
           dibuat" padahal tidak ada akun baru yang tercipta. */
        const alreadyRegistered = !!(data && data.user && Array.isArray(data.user.identities) && data.user.identities.length === 0);
        if (alreadyRegistered) {
          throw new Error('Email ini sudah terdaftar sebagai akun admin. Gunakan email lain, atau ubah perannya lewat tabel "Kelola Pengguna & Peran".');
        }

        /* Simpan peran & status aktif akun baru ke admin_profiles supaya langsung tampil
           benar di tabel "Kelola Pengguna & Peran" tanpa menunggu akun tsb login sendiri. */
        if (data && data.user && data.user.id) {
          try { await DataLayer.createProfile(data.user.id, email, role); }
          catch (profileErr) { console.warn('Gagal menyimpan profil peran akun baru:', profileErr.message); }
        }
        await DataLayer.logAudit('create_account', email, role === 'admin' ? 'Admin' : 'Operator');

        /* Klien kedua ini tidak menyimpan sesi (persistSession:false), tapi tetap keluarkan
           sesi in-memory-nya untuk berjaga-jaga. */
        await sbForNewAdmin.auth.signOut().catch(() => { });

        const needsConfirmation = !!(data && data.user && !data.session);
        document.getElementById('ta-email').value = '';
        document.getElementById('ta-password').value = '';
        document.getElementById('ta-confirm-password').value = '';
        document.getElementById('ta-role').value = 'admin';
        msg.innerHTML = `<div class="form-msg ok">Akun ${role === 'admin' ? 'admin' : 'operator'} baru (${escapeHtml(email)}) berhasil dibuat.${needsConfirmation ? ' Akun ini perlu mengklik tautan konfirmasi di emailnya sebelum bisa masuk (kecuali opsi "Confirm email" sudah dimatikan di Supabase).' : ' Akun sudah bisa langsung dipakai untuk masuk.'}</div>`;
        await renderPenggunaTable();
        renderAuditLog();
      } catch (err) {
        msg.innerHTML = `<div class="form-msg err">Gagal membuat akun: ${escapeHtml(errMsg(err))}</div>`;
      } finally {
        btn.disabled = false;
      }
    });

    /* =====================================================================
       INIT
       ===================================================================== */
    function syncMobileMenuButton() {
      document.getElementById('btn-menu-mobile').classList.toggle('hidden', window.innerWidth >= 900);
    }
    window.addEventListener('resize', syncMobileMenuButton);

    (async function init() {
      await initSession();
      await router();
      syncMobileMenuButton();
    })();
  </script>
</body>

</html>

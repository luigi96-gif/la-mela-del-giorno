(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var meta = null;
  var state = { running: false, phase: "", busy: [] };
  var seenMessages = {};
  var currentDate = null; // null = ultima edizione

  var MONTHS = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
    "agosto", "settembre", "ottobre", "novembre", "dicembre"];
  var DAYS = ["Domenica", "Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato"];
  var TYPE_LABEL = { notizia: "Notizia", indiscrezione: "Indiscrezione", analisi: "Analisi" };
  var CHECK_LABEL = { confermata: "Confermata", indiscrezione: "Voce non confermata", da_verificare: "Da verificare" };

  function el(tag, cls, text) {
    var node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined && text !== null) node.textContent = text;
    return node;
  }

  function formatDate(iso) {
    var p = iso.split("-");
    return parseInt(p[2], 10) + " " + MONTHS[parseInt(p[1], 10) - 1] + " " + p[0];
  }

  // ---------- versione da leggere ovunque: contenuti cifrati, si aprono con la password ----------

  var STATIC = !!window.__GIORNALE_STATICO__;
  var secretKey = null;

  function b64(text) {
    var bin = atob(text), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function openFile(path) {
    return fetch(path, { cache: "no-cache" }).then(function (res) {
      if (!res.ok) throw new Error("File non trovato");
      return res.json();
    }).then(function (box) {
      return crypto.subtle.decrypt({ name: "AES-GCM", iv: b64(box.iv) }, secretKey, b64(box.data));
    }).then(function (plain) {
      return JSON.parse(new TextDecoder().decode(plain));
    });
  }

  function staticApi(path) {
    if (path === "/api/meta") return openFile("dati/meta.json");
    if (path.indexOf("/api/edition") === 0 && path.indexOf("/api/edition/") !== 0) {
      var wanted = path.match(/date=([0-9-]+)/);
      return openFile("dati/edizioni.json").then(function (index) {
        var date = wanted ? wanted[1] : index.dates[0];
        if (!date) return { edition: null, dates: index.dates };
        return openFile("dati/edizione-" + date + ".json").then(function (edition) {
          return { edition: edition, dates: index.dates };
        });
      });
    }
    return Promise.reject(new Error("Non disponibile nella versione da leggere"));
  }

  function unlock(password) {
    return fetch("dati/chiave.json", { cache: "no-cache" }).then(function (r) { return r.json(); }).then(function (info) {
      return crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]).then(function (base) {
        return crypto.subtle.deriveKey({ name: "PBKDF2", salt: b64(info.salt), iterations: info.iterations, hash: "SHA-256" },
          base, { name: "AES-GCM", length: 256 }, true, ["decrypt"]);
      });
    }).then(function (key) {
      secretKey = key;
      return openFile("dati/meta.json").then(function () {
        // ricordata su questo dispositivo, così non la si riscrive ogni volta
        return crypto.subtle.exportKey("raw", key).then(function (raw) {
          try { localStorage.setItem("chiave", btoa(String.fromCharCode.apply(null, new Uint8Array(raw)))); } catch (e) {}
        });
      });
    });
  }

  function restoreKey() {
    var saved = null;
    try { saved = localStorage.getItem("chiave"); } catch (e) {}
    if (!saved) return Promise.resolve(false);
    return crypto.subtle.importKey("raw", b64(saved), { name: "AES-GCM" }, true, ["decrypt"]).then(function (key) {
      secretKey = key;
      return openFile("dati/meta.json").then(function () { return true; });
    }).catch(function () {
      try { localStorage.removeItem("chiave"); } catch (e) {}
      return false;
    });
  }

  function api(path, options) {
    if (STATIC) return staticApi(path);
    options = options || {};
    if (options.body) {
      options.headers = { "Content-Type": "application/json" };
      options.body = JSON.stringify(options.body);
    }
    return fetch(path, options).then(function (res) {
      if (res.status === 401) { showLogin(); throw new Error("Accesso richiesto"); }
      return res.json().then(function (data) {
        if (!res.ok) throw new Error(data.detail || "Errore");
        return data;
      });
    });
  }

  // ---------- avvio ----------

  function init() {
    if (STATIC) {
      fetch("dati/chiave.json", { cache: "no-cache" }).then(function (r) { return r.json(); }).then(function (info) {
        $("paper-name").textContent = info.nome;
        document.title = info.nome;
      }).catch(function () {});
      restoreKey().then(function (ok) {
        if (ok) showPaper(); else showLogin();
      });
      return;
    }
    showPaper();
  }

  function showPaper() {
    api("/api/meta").then(function (m) {
      $("login").hidden = true;
      meta = m;
      document.title = m.name;
      $("paper-name").textContent = m.name;
      $("paper-tagline").textContent = m.tagline;
      $("edition-time").textContent = m.edition_time;
      $("demo-banner").hidden = !m.demo;
      var now = new Date();
      $("today").textContent = DAYS[now.getDay()] + " " + now.getDate() + " " + MONTHS[now.getMonth()] + " " + now.getFullYear();
      if (m.reader) {
        // Da un altro dispositivo della rete di casa: solo il giornale, in lettura.
        document.body.classList.add("reader");
        renderSectionBar();
        loadEdition();
      } else if (m.auth_required && !m.authed) { showLogin(); } else { start(); }
    });
  }

  function start() {
    $("login").hidden = true;
    renderSectionBar();
    renderAgents();
    Office.init($("office"), meta, openAgentChat);
    loadEdition();
    loadRoom();
    loadSettings();
    connectEvents();
  }

  function showLogin() { $("login").hidden = false; $("password").focus(); }

  $("login-form").addEventListener("submit", function (e) {
    e.preventDefault();
    $("login-msg").textContent = "";
    if (STATIC) {
      $("login-msg").textContent = "Apro il giornale…";
      unlock($("password").value).then(function () {
        $("password").value = "";
        $("login-msg").textContent = "";
        showPaper();
      }).catch(function () { $("login-msg").textContent = "Password errata."; });
      return;
    }
    fetch("/api/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password: $("password").value })
    }).then(function (res) {
      if (res.ok) { $("password").value = ""; start(); }
      else { $("login-msg").textContent = "Password errata."; }
    });
  });

  // ---------- schede ----------

  Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (tab) {
    tab.addEventListener("click", function () { showView(tab.dataset.view); });
  });

  function showView(name) {
    Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (t) {
      t.classList.toggle("active", t.dataset.view === name);
    });
    ["edizione", "redazione", "impostazioni"].forEach(function (v) {
      $("view-" + v).hidden = v !== name;
    });
    if (name === "redazione") scrollMessages(true);
  }

  // ---------- edizione ----------

  function loadEdition() {
    var query = currentDate ? "?date=" + encodeURIComponent(currentDate) : "";
    return api("/api/edition" + query).then(function (data) {
      lastEdition = data.edition;
      lastDates = data.dates;
      if (!currentDate && !meta.reader) {
        Office.setEdition(data.edition);
        loadFrontPages(data.dates);
      }
      renderArchive(data.dates, data.edition);
      renderEdition(data.edition);
    }).catch(function () {});
  }

  // Le ultime prime pagine pubblicate, appese al muro dell'ufficio.
  function loadFrontPages(dates) {
    Promise.all(dates.slice(0, 5).map(function (d) {
      return api("/api/edition?date=" + encodeURIComponent(d)).then(function (r) { return r.edition; });
    })).then(function (list) {
      Office.setFrontPages(list.filter(function (e) { return e && e.stato !== "bozza"; }).slice(0, 3).map(function (e) {
        return { date: e.date, numero: e.numero, titolo: e.stories[0].titolo, occhiello: e.stories[0].occhiello };
      }));
    }).catch(function () {});
  }

  function renderArchive(dates, edition) {
    var select = $("archive");
    select.textContent = "";
    if (!dates.length) { select.appendChild(new Option("Nessuna edizione", "")); return; }
    dates.forEach(function (d) { select.appendChild(new Option(formatDate(d), d)); });
    if (edition) select.value = edition.date;
  }

  $("archive").addEventListener("change", function (e) {
    currentDate = e.target.value || null;
    loadEdition();
  });

  function renderEdition(edition) {
    var root = $("edition");
    root.textContent = "";
    if (!edition) {
      $("edition-number").textContent = "";
      $("edition-info").textContent = "";
      var empty = el("div", "empty");
      empty.appendChild(el("h2", null, "La prima edizione non è ancora uscita"));
      empty.appendChild(el("p", "muted", "Vai in Redazione e chiedi di far uscire l'edizione: potrai seguire il lavoro in diretta."));
      var go = el("button", "primary", "Vai in redazione");
      go.addEventListener("click", function () { showView("redazione"); });
      empty.appendChild(go);
      root.appendChild(empty);
      return;
    }
    $("edition-number").textContent = "Edizione n. " + edition.numero + " · " + formatDate(edition.date);
    var closed = new Date(edition.generated_at * 1000);
    var info = edition.stories.length + " pezzi da " + edition.articoli_letti +
      " notizie lette su " + edition.fonti_lette + " fonti · chiusa alle " +
      closed.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" }) +
      (edition.demo ? " · dimostrativa" : "");
    if (edition.consumo && edition.consumo.richieste && !meta.reader) {
      var tokens = edition.consumo.token_letti + edition.consumo.token_scritti;
      info += " · " + edition.consumo.richieste + " richieste a Claude, circa " +
        Math.round(tokens / 1000).toLocaleString("it-IT") + " mila token";
    }
    $("edition-info").textContent = info;

    if (edition.stato === "bozza") {
      var bar = el("div", "draft-bar");
      var text = el("div");
      text.appendChild(el("b", null, "Bozza in attesa del tuo via. "));
      text.appendChild(document.createTextNode("Leggila, chiedi di rifare o togli i pezzi che non ti convincono, poi pubblica."));
      bar.appendChild(text);
      var publish = el("button", "primary", "Pubblica l'edizione");
      publish.addEventListener("click", function () {
        publish.disabled = true;
        api("/api/edition/publish", { method: "POST", body: { date: edition.date } })
          .catch(function (err) { publish.disabled = false; alert(err.message); });
      });
      bar.appendChild(publish);
      root.appendChild(bar);
    }

    // si vede solo la sezione scelta nel selettore in fondo alla pagina
    var stories = edition.stories.filter(function (s) { return sectionOf(s) === currentSection(); });
    if (!stories.length) {
      var none = el("div", "empty");
      none.appendChild(el("h2", null, "Nessun pezzo in questa sezione"));
      none.appendChild(el("p", "muted", "Oggi la redazione non ha pubblicato pezzi per «" + sectionTitle(currentSection()) + "»."));
      root.appendChild(none);
      return;
    }
    if (edition.incompleta) {
      root.appendChild(el("div", "incomplete-note", edition.in_arrivo
        ? "Edizione in completamento: altri " + edition.in_arrivo + " pezzi arrivano appena si azzera il limite d'uso dell'abbonamento."
        : "Edizione in completamento: Il punto e l'impaginazione arrivano appena si azzera il limite d'uso dell'abbonamento."));
    }
    // le cose da sapere oggi, in trenta secondi
    if (edition.in_breve && edition.in_breve.length) {
      var brief = el("section", "brief");
      brief.appendChild(el("h3", null, "In breve"));
      var list = el("ol");
      edition.in_breve.forEach(function (line) { list.appendChild(el("li", null, line)); });
      brief.appendChild(list);
      root.appendChild(brief);
    }
    var front = el("div", "front");
    front.appendChild(renderStory(stories[0], true, edition));
    var editorial = el("aside", "editorial");
    editorial.appendChild(el("h3", null, "Il punto"));
    if (edition.editoriale.testo) {
      editorial.appendChild(el("h4", null, edition.editoriale.titolo));
      paragraphs(edition.editoriale.testo).forEach(function (p) { editorial.appendChild(el("p", null, p)); });
    } else {
      editorial.appendChild(el("p", "muted", "Arriva a edizione completa."));
    }
    editorial.appendChild(el("p", "sign", "— " + meta.agents.direttrice.name + ", " + meta.agents.direttrice.role.toLowerCase()));
    front.appendChild(editorial);
    root.appendChild(front);

    var grid = el("div", "grid");
    stories.slice(1).forEach(function (s) { grid.appendChild(renderStory(s, false, edition)); });
    root.appendChild(grid);
  }

  // ---------- ascolta l'edizione ----------

  var listening = false;

  function setListening(on) {
    listening = on;
    $("listen").textContent = on ? "Ferma" : "Ascolta";
    $("listen").setAttribute("aria-pressed", on ? "true" : "false");
  }

  function stopListening() {
    if (listening && window.speechSynthesis) speechSynthesis.cancel();
    setListening(false);
  }

  // frasi brevi: alcuni browser interrompono le letture troppo lunghe
  function sentences(text) {
    return (text || "").replace(/\s+/g, " ").match(/[^.!?…]+[.!?…»"]*\s*/g) || [];
  }

  function listenParts() {
    var edition = lastEdition, parts = [];
    if (!edition) return parts;
    if (edition.in_breve && edition.in_breve.length) {
      parts.push("In breve.");
      edition.in_breve.forEach(function (line) { parts.push(line); });
    }
    edition.stories.filter(function (s) { return sectionOf(s) === currentSection(); }).forEach(function (s) {
      parts.push(s.titolo + ".");
      sentences(s.corpo).forEach(function (t) { if (t.trim()) parts.push(t.trim()); });
    });
    return parts;
  }

  $("listen").addEventListener("click", function () {
    if (!("speechSynthesis" in window)) { alert("Questo browser non sa leggere ad alta voce."); return; }
    if (listening) { stopListening(); return; }
    var parts = listenParts();
    if (!parts.length) return;
    var voices = speechSynthesis.getVoices();
    var voice = voices.filter(function (v) { return /^it/i.test(v.lang) && /premium|enhanced|alice|federica|luca/i.test(v.name); })[0]
      || voices.filter(function (v) { return /^it/i.test(v.lang); })[0];
    speechSynthesis.cancel();
    parts.forEach(function (text, i) {
      var u = new SpeechSynthesisUtterance(text);
      u.lang = "it-IT";
      if (voice) u.voice = voice;
      if (i === parts.length - 1) u.onend = function () { setListening(false); };
      speechSynthesis.speak(u);
    });
    setListening(true);
  });

  // ---------- ricerca nell'archivio ----------

  var lastDates = [];
  var archiveCache = {};

  function fold(text) {
    return (text || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  function loadArchive() {
    return Promise.all(lastDates.map(function (d) {
      if (archiveCache[d]) return archiveCache[d];
      return api("/api/edition?date=" + encodeURIComponent(d)).then(function (r) { archiveCache[d] = r.edition; return r.edition; });
    }));
  }

  function closeSearch() {
    $("search-results").hidden = true;
    $("edition").hidden = false;
  }

  $("search-form").addEventListener("submit", function (e) {
    e.preventDefault();
    var query = $("search").value.trim();
    var box = $("search-results");
    if (query.length < 2) { closeSearch(); return; }
    stopListening();
    box.hidden = false;
    $("edition").hidden = true;
    box.textContent = "";
    box.appendChild(el("p", "muted", "Cerco in " + lastDates.length + " edizioni…"));
    var words = fold(query).split(/\s+/).filter(Boolean);
    loadArchive().then(function (editions) {
      var hits = [];
      editions.forEach(function (edition) {
        if (!edition || (meta.reader && edition.stato === "bozza")) return;
        edition.stories.forEach(function (s) {
          var hay = fold([s.titolo, s.occhiello, s.sommario, s.corpo].join(" "));
          if (words.every(function (w) { return hay.indexOf(w) !== -1; })) hits.push({ date: edition.date, story: s });
        });
      });
      box.textContent = "";
      var head = el("div", "search-head");
      head.appendChild(el("b", null, hits.length ? hits.length + (hits.length === 1 ? " pezzo" : " pezzi") + " per «" + query + "»" : "Nessun pezzo per «" + query + "»"));
      var back = el("button", "tool-btn", "Torna all'edizione");
      back.type = "button";
      back.addEventListener("click", function () { $("search").value = ""; closeSearch(); });
      head.appendChild(back);
      box.appendChild(head);
      hits.forEach(function (hit) {
        var item = el("button", "search-hit");
        item.type = "button";
        item.appendChild(el("span", "muted small", formatDate(hit.date) + " · " + sectionTitle(sectionOf(hit.story)) + " · " + hit.story.autore));
        item.appendChild(el("strong", null, hit.story.titolo));
        item.appendChild(el("span", null, hit.story.sommario));
        item.addEventListener("click", function () { openStory(hit.date, hit.story); });
        box.appendChild(item);
      });
    }).catch(function () { box.textContent = "Ricerca non riuscita."; });
  });

  function openStory(date, story) {
    try { localStorage.setItem("sezione", sectionOf(story)); } catch (e) {}
    renderSectionBar();
    closeSearch();
    currentDate = date;
    loadEdition().then(function () {
      var node = $("pezzo-" + story.id);
      if (!node) return;
      var more = node.querySelector(".more");
      if (more && node.querySelector(".body").hidden) more.click();
      node.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  // ---------- selettore delle sezioni ----------

  var lastEdition = null;

  function sectionKeys() { return Object.keys(meta.sections || {}); }

  function sectionTitle(key) { return meta.sections && meta.sections[key] ? meta.sections[key].title : key; }

  function currentSection() {
    var saved = null;
    try { saved = localStorage.getItem("sezione"); } catch (e) {}
    var keys = sectionKeys();
    return keys.indexOf(saved) !== -1 ? saved : keys[0];
  }

  function chooseSection(key) {
    try { localStorage.setItem("sezione", key); } catch (e) {}
    stopListening();
    closeSearch();
    renderSectionBar();
    showView("edizione");
    renderEdition(lastEdition);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function renderSectionBar() {
    var bar = $("section-bar");
    var keys = sectionKeys();
    bar.textContent = "";
    bar.hidden = keys.length < 2;
    var active = currentSection();
    keys.forEach(function (key) {
      var b = el("button", "section-btn" + (key === active ? " active" : ""), sectionTitle(key));
      b.type = "button";
      b.setAttribute("aria-pressed", key === active ? "true" : "false");
      b.addEventListener("click", function () { chooseSection(key); });
      bar.appendChild(b);
    });
    document.body.classList.toggle("with-sections", !bar.hidden);
  }

  function sectionOf(story) {
    if (story.sezione) return story.sezione;
    var sections = meta.sections || {};
    var found = "apple";
    Object.keys(sections).forEach(function (key) {
      if (sections[key].desks.indexOf(story.desk) !== -1) found = key;
    });
    return found;
  }

  function paragraphs(text) {
    return (text || "").split(/\n\s*\n/).map(function (p) { return p.trim(); }).filter(Boolean);
  }

  function storyActions(story, edition) {
    var box = el("div", "story-actions");
    var redo = el("button", null, "Chiedi di rifarlo");
    var remove = el("button", null, "Togli dall'edizione");
    var form = el("form", "redo-form");
    form.hidden = true;
    var note = el("textarea");
    note.rows = 2;
    note.placeholder = "Che cosa deve cambiare " + story.autore + "? (es. più breve, più contesto sui prezzi…)";
    var send = el("button", "primary", "Manda a " + story.autore);
    send.type = "submit";
    form.appendChild(note); form.appendChild(send);
    var msg = el("div", "muted small");
    redo.addEventListener("click", function () { form.hidden = !form.hidden; if (!form.hidden) note.focus(); });
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      msg.textContent = "";
      api("/api/edition/story/redo", { method: "POST", body: { date: edition.date, story_id: story.id, note: note.value } })
        .then(function () { showView("redazione"); })
        .catch(function (err) { msg.textContent = err.message; });
    });
    remove.addEventListener("click", function () {
      if (!confirm("Togliere «" + story.titolo + "» dall'edizione?")) return;
      api("/api/edition/story/remove", { method: "POST", body: { date: edition.date, story_id: story.id } })
        .catch(function (err) { msg.textContent = err.message; });
    });
    if (story.articoli && story.articoli.length) box.appendChild(redo);
    box.appendChild(remove);
    box.appendChild(form);
    box.appendChild(msg);
    return box;
  }

  function renderStory(story, lead, edition) {
    var article = el("article", "story" + (lead ? " lead" : ""));
    article.id = "pezzo-" + story.id;
    article.appendChild(el("div", "kicker", story.occhiello));
    article.appendChild(el("h2", null, story.titolo));
    article.appendChild(el("p", "summary", story.sommario));

    var byline = el("div", "byline");
    byline.appendChild(el("span", null, "di " + story.autore));
    byline.appendChild(el("span", null, "· " + (TYPE_LABEL[story.tipo] || story.tipo)));
    byline.appendChild(el("span", "badge " + story.verifica.esito, CHECK_LABEL[story.verifica.esito] || story.verifica.esito));
    if (story.su_richiesta) byline.appendChild(el("span", "badge richiesta", "Su tua richiesta"));
    article.appendChild(byline);

    var body = el("div", "body");
    body.hidden = !lead;
    paragraphs(story.corpo).forEach(function (p) { body.appendChild(el("p", null, p)); });
    var note = el("div", "check-note");
    note.appendChild(el("b", null, "Verifica: "));
    note.appendChild(document.createTextNode(story.verifica.nota));
    body.appendChild(note);
    var list = el("ul", "sources");
    story.fonti.forEach(function (f) {
      var li = el("li");
      var a = el("a", null, f.titolo);
      a.href = f.url; a.target = "_blank"; a.rel = "noopener noreferrer";
      li.appendChild(document.createTextNode(f.nome + " — "));
      li.appendChild(a);
      list.appendChild(li);
    });
    body.appendChild(list);
    if (!meta.reader) body.appendChild(storyActions(story, edition));

    if (!lead) {
      var more = el("button", "more", "Leggi il pezzo");
      more.addEventListener("click", function () {
        body.hidden = !body.hidden;
        more.textContent = body.hidden ? "Leggi il pezzo" : "Chiudi";
      });
      article.appendChild(more);
    }
    article.appendChild(body);
    return article;
  }

  // ---------- stanza ----------

  function avatar(agentId) {
    var agent = meta.agents[agentId];
    var node = el("span", "avatar", agent ? agent.name.charAt(0) : "Tu");
    node.style.background = agent ? agent.color : "#555";
    return node;
  }

  function renderAgents() {
    var root = $("agents");
    root.textContent = "";
    Object.keys(meta.agents).forEach(function (id) {
      var agent = meta.agents[id];
      var button = el("button", "agent");
      button.type = "button";
      button.id = "agent-" + id;
      button.title = agent.bio;
      button.appendChild(avatar(id));
      var who = el("span", "who");
      who.appendChild(el("div", "name", agent.name));
      who.appendChild(el("div", "role", agent.role));
      button.appendChild(who);
      button.appendChild(el("span", "status"));
      button.addEventListener("click", function () { openAgentChat(id); });
      root.appendChild(button);
    });
  }

  // ---------- conversazione a tu per tu con un agente ----------

  var allMessages = [];
  var chatAgent = null;
  var chatPending = false;

  function inThread(m, id) {
    if (m.author === id && m.kind !== "system") return true;
    if (m.author !== "editore") return false;
    var text = m.text.toLowerCase();
    return text.indexOf("@" + meta.agents[id].name.toLowerCase()) !== -1 || /@(tutti|redazione)\b/.test(text);
  }

  function threadBubble(m) {
    var mine = m.author === "editore";
    var text = mine ? m.text.replace(new RegExp("^@" + meta.agents[chatAgent].name + "\\s*", "i"), "") : m.text;
    return el("div", "ac-msg" + (mine ? " me" : ""), text);
  }

  function chatStatus() {
    if (!chatAgent) return;
    var busy = state.busy.indexOf(chatAgent) !== -1;
    var name = meta.agents[chatAgent].name;
    $("ac-status").textContent = chatPending && busy ? "sta scrivendo la risposta…"
      : busy ? "sta lavorando a un pezzo, ma ti ascolta"
      : meta.agents[chatAgent].role + " · in attesa di istruzioni";
    $("ac-input").placeholder = "Scrivi a " + name + "…";
  }

  function renderThread() {
    var box = $("ac-thread");
    box.textContent = "";
    var items = allMessages.filter(function (m) { return inThread(m, chatAgent); }).slice(-40);
    if (!items.length) {
      box.appendChild(el("div", "ac-empty", meta.agents[chatAgent].name + " ti guarda, in attesa di istruzioni."));
    }
    items.forEach(function (m) { box.appendChild(threadBubble(m)); });
    box.scrollTop = box.scrollHeight;
  }

  function openAgentChat(id) {
    chatAgent = id;
    chatPending = false;
    var agent = meta.agents[id];
    $("ac-avatar").textContent = agent.name.charAt(0);
    $("ac-avatar").style.background = agent.color;
    $("ac-name").textContent = agent.name;
    $("agent-chat").hidden = false;
    Office.select(id);
    chatStatus();
    renderThread();
    $("office").scrollIntoView({ block: "nearest", behavior: "smooth" });
    $("ac-input").focus({ preventScroll: true });
  }

  function closeAgentChat() {
    chatAgent = null;
    $("agent-chat").hidden = true;
    Office.select(null);
  }

  $("ac-close").addEventListener("click", closeAgentChat);
  document.addEventListener("keydown", function (e) { if (e.key === "Escape" && chatAgent) closeAgentChat(); });

  $("ac-form").addEventListener("submit", function (e) {
    e.preventDefault();
    var input = $("ac-input");
    var text = input.value.trim();
    if (!text || !chatAgent) return;
    input.value = "";
    input.style.height = "auto";
    chatPending = true;
    chatStatus();
    api("/api/room/message", { method: "POST", body: { text: text, to: chatAgent } })
      .catch(function () { input.value = text; chatPending = false; chatStatus(); });
  });
  $("ac-input").addEventListener("keydown", function (e) {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); $("ac-form").requestSubmit(); }
  });
  $("ac-input").addEventListener("input", function (e) {
    e.target.style.height = "auto";
    e.target.style.height = Math.min(e.target.scrollHeight, 110) + "px";
  });

  function loadRoom() {
    api("/api/room").then(function (data) {
      $("messages").textContent = "";
      seenMessages = {};
      allMessages = [];
      $("ac-thread").textContent = "";
      data.messages.forEach(addMessage);
      applyState(data.state);
      Office.setBoard(data.board || []);
      scrollMessages(true);
    }).catch(function () {});
  }

  function addMessage(m) {
    if (seenMessages[m.id]) return;
    seenMessages[m.id] = true;
    allMessages.push(m);
    if (chatAgent && inThread(m, chatAgent)) {
      if (m.author === chatAgent) chatPending = false;
      var thread = $("ac-thread");
      var hint = thread.querySelector(".ac-empty");
      if (hint) hint.remove();
      thread.appendChild(threadBubble(m));
      thread.scrollTop = thread.scrollHeight;
      chatStatus();
    }
    var box = $("messages");
    var stick = box.scrollHeight - box.scrollTop - box.clientHeight < 80;
    var time = new Date(m.ts * 1000).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
    var row;
    if (m.kind === "system" || !(m.author === "editore" || meta.agents[m.author])) {
      row = el("div", "msg system");
      row.appendChild(el("div", "text", m.text));
    } else if (m.author === "editore") {
      row = el("div", "msg me");
      var mine = el("div", "bubble");
      mine.appendChild(el("div", "text", m.text));
      row.appendChild(mine);
    } else {
      var agent = meta.agents[m.author];
      row = el("div", "msg");
      row.appendChild(avatar(m.author));
      var bubble = el("div", "bubble");
      var head = el("div", "head");
      head.appendChild(el("b", null, agent.name));
      head.appendChild(el("span", "muted", " · " + agent.role + " · " + time));
      bubble.appendChild(head);
      bubble.appendChild(el("div", "text", m.text));
      row.appendChild(bubble);
    }
    box.appendChild(row);
    if (stick) scrollMessages(true);
  }

  function scrollMessages(force) {
    var box = $("messages");
    if (force) box.scrollTop = box.scrollHeight;
  }

  function applyState(s) {
    state = s;
    Office.setState(s);
    chatStatus();
    $("live-dot").hidden = !s.running;
    $("phase").textContent = s.running ? "In lavorazione: " + s.phase : "Redazione a riposo";
    $("run").disabled = s.running;
    $("run").textContent = s.running ? "Edizione in lavorazione…" : "Fai uscire l'edizione ora";
    Object.keys(meta.agents).forEach(function (id) {
      var node = $("agent-" + id);
      var busy = s.busy.indexOf(id) !== -1;
      node.classList.toggle("busy", busy);
      node.querySelector(".status").textContent = busy ? "al lavoro" : "";
    });
    var names = s.busy.map(function (id) { return meta.agents[id] ? meta.agents[id].name : id; });
    $("typing").textContent = names.length
      ? names.join(", ") + (names.length > 1 ? " stanno lavorando…" : " sta lavorando…")
      : "";
  }

  $("run").addEventListener("click", function () {
    $("run").disabled = true;
    api("/api/edition/run", { method: "POST" }).catch(function () { $("run").disabled = false; });
  });

  $("composer").addEventListener("submit", function (e) {
    e.preventDefault();
    var input = $("input");
    var text = input.value.trim();
    if (!text) return;
    input.value = "";
    input.style.height = "auto";
    api("/api/room/message", { method: "POST", body: { text: text } }).catch(function () { input.value = text; });
  });

  $("input").addEventListener("keydown", function (e) {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      $("composer").requestSubmit();
    }
  });
  $("input").addEventListener("input", function (e) {
    e.target.style.height = "auto";
    e.target.style.height = Math.min(e.target.scrollHeight, 140) + "px";
  });

  function connectEvents() {
    var source = new EventSource("/api/events");
    source.addEventListener("message", function (e) {
      var m = JSON.parse(e.data);
      addMessage(m);
      if (meta.agents[m.author]) Office.say(m.author, m.text);
    });
    source.addEventListener("board", function (e) { Office.setBoard(JSON.parse(e.data)); });
    source.addEventListener("status", function (e) { applyState(JSON.parse(e.data)); });
    source.addEventListener("edition", function (e) {
      var info = JSON.parse(e.data);
      // prima la scena (la stampante), poi il nuovo contenuto: così la prima pagina arriva al muro con Elena
      if (info.evento === "pubblicata" || info.evento === "bozza") Office.printEdition(info.evento);
      currentDate = null;
      loadEdition();
    });
    source.addEventListener("directives", function (e) { renderDirectives(JSON.parse(e.data)); });
    // Dopo una disconnessione il browser si ricollega da solo: riallinea lo storico.
    source.addEventListener("open", function () { loadRoom(); });
  }

  // ---------- impostazioni ----------

  function loadSettings() {
    api("/api/settings").then(function (data) {
      renderSources(data.sources);
      renderDirectives(data.directives);
      $("cloud-box").hidden = !meta.cloud;
      $("lan-info").textContent = meta.lan_url
        ? "Dal telefono o dal tablet, sulla stessa rete Wi-Fi di questo Mac, apri " + meta.lan_url +
          " : vedrai il giornale in sola lettura (solo le edizioni pubblicate), finché l'app è accesa."
        : "La lettura da altri dispositivi non è disponibile con questa configurazione.";
    }).catch(function () {});
  }

  function renderSources(sources) {
    var list = $("sources");
    list.textContent = "";
    sources.forEach(function (s) {
      var li = el("li");
      var info = el("div");
      info.appendChild(el("div", null, s.name + (s.lang === "it" ? " · italiano" : " · inglese")));
      info.appendChild(el("div", "url", s.url));
      li.appendChild(info);
      var remove = el("button", null, "Togli");
      remove.addEventListener("click", function () {
        api("/api/sources?url=" + encodeURIComponent(s.url), { method: "DELETE" })
          .then(function (d) { renderSources(d.sources); });
      });
      li.appendChild(remove);
      list.appendChild(li);
    });
  }

  function renderDirectives(directives) {
    var list = $("directives");
    list.textContent = "";
    if (!directives.length) {
      list.appendChild(el("li", "muted", "Nessuna indicazione per ora."));
      return;
    }
    directives.forEach(function (d) {
      var li = el("li");
      li.appendChild(el("div", null, d.text));
      var remove = el("button", null, "Togli");
      remove.addEventListener("click", function () {
        api("/api/directives/" + d.id, { method: "DELETE" })
          .then(function (res) { renderDirectives(res.directives); });
      });
      li.appendChild(remove);
      list.appendChild(li);
    });
  }

  $("cloud-push").addEventListener("click", function () {
    var msg = $("cloud-msg");
    msg.textContent = "Invio in corso…";
    api("/api/cloud/push", { method: "POST" })
      .then(function (d) { msg.textContent = d.detail; })
      .catch(function (err) { msg.textContent = err.message; });
  });


  $("source-form").addEventListener("submit", function (e) {
    e.preventDefault();
    var msg = $("source-msg");
    msg.textContent = "Controllo il feed…";
    api("/api/sources", {
      method: "POST",
      body: { name: $("source-name").value, url: $("source-url").value, lang: $("source-lang").value }
    }).then(function (d) {
      renderSources(d.sources);
      msg.textContent = "Fonte aggiunta (" + d.detail + ").";
      $("source-name").value = "";
      $("source-url").value = "";
    }).catch(function (err) { msg.textContent = err.message; });
  });

  init();
})();

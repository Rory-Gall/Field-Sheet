/* KEI Field Sheet - the sender (build 22).

   ONE copy of the code that sends finished sheets. The service worker loads it
   (importScripts) and does all the sending, so a sheet goes out when the phone gets signal
   even with the app closed: Android wakes the worker for a "background sync". The page loads
   the same file and only sends by itself when there is no worker (a ?nosw preview, or a
   browser without one) - there is never more than one sender.

   Where a finished sheet goes, when there is signal:
     1. _KEI Monitoring Sites\<site>\Field Sheets\  (the folder must already exist)
     2. Hydro Processor\Sites\_Staged\<station>\ + <file>.stage.json (kei-stage/1), only when
        Sites\<station>.keisite.json exists - the Processor's "Approve the latest uploads".
   Build 33: before either, a small notice <draftId>.waiting.json (kei-waiting/1) goes to that same
   _Staged\<station>\ folder, so the office sees a sheet is on its way while the big file is still
   uploading; it carries the progress and is taken away once both uploads are done.
   Folders are reached by their OneDrive ids: a renamed folder makes the send fail loudly,
   never misfile. Sign-in is the KEI Board's "KEI Task Board" Microsoft app (delegated
   Sites.ReadWrite.All: a person can only write where they already can). Microsoft caps a
   phone web app's sign-in at 24 hours; a sheet that can't send says why and waits.

   Everything lives in IndexedDB "kei-fieldsheet" / store "photos", the store the page already
   uses, so page and worker read the same state:
     meta:auth        the sign-in (tokens, name, email, signedInAt)
     ob:<draftId>     one outbox entry per finished sheet (its state and progress)
     outbox:<draftId> the finished file itself
   Each entry is changed inside a single transaction, never as a whole list. */
(function (G) {
  "use strict";

  var CFG = {
    tenantId: "ff85c429-ed5b-458f-aca2-478fd1f538e1",
    clientId: "d6fb3bfc-b421-4d95-b7b3-f3cce4dad884",
    driveId: "b!J072BbRI40OmbEK7ALh-5w9sPauU7UtDrafivbxV73LAogmnPVCrQoijSVCwNloy",
    monSites: "01REBDVIUZKO6DNSV6QFGJHANVCNZU3ODA",
    staged: "01REBDVIRGPEVDFQVXCJBIWYQXVUJ26ITW",
    graph: "https://graph.microsoft.com/v1.0"
  };
  var AUTHORITY = "https://login.microsoftonline.com/" + CFG.tenantId + "/oauth2/v2.0";
  var SCOPES = "https://graph.microsoft.com/Sites.ReadWrite.All https://graph.microsoft.com/User.Read offline_access openid profile";
  var SIGNED_OUT = "signed-out";
  /* the phone's station name -> the Processor's station file, where they differ */
  var PROC_ID = { "CoN_MtnStn Resevoir": "CoN_Mtn Station Reservoir" };
  var IN_WORKER = typeof WorkerGlobalScope !== "undefined" && G instanceof WorkerGlobalScope;

  /* ---------- storage ---------- */
  var DB = null;
  function db() {
    if (DB) return Promise.resolve(DB);
    return new Promise(function (res, rej) {
      var r = indexedDB.open("kei-fieldsheet", 1);
      r.onupgradeneeded = function () { r.result.createObjectStore("photos"); };
      r.onsuccess = function () { DB = r.result; DB.onversionchange = function () { DB.close(); DB = null; }; res(DB); };
      r.onerror = function () { rej(r.error || new Error("storage unavailable")); };
    });
  }
  function op(mode, fn) {
    return db().then(function (d) {
      return new Promise(function (res, rej) {
        var tx = d.transaction("photos", mode), st = tx.objectStore("photos");
        var out = fn(st);
        tx.oncomplete = function () { res(out && "result" in out ? out.result : undefined); };
        tx.onerror = function () { rej(tx.error); };
        tx.onabort = function () { rej(tx.error || new Error("storage write was cancelled")); };
      });
    });
  }
  function get(k) { return op("readonly", function (st) { return st.get(k); }); }
  function put(k, v) { return op("readwrite", function (st) { return st.put(v, k); }); }
  function del(k) { return op("readwrite", function (st) { return st.delete(k); }); }

  /* tell every open copy of the app that something changed, so it redraws */
  function notify(kind) {
    if (IN_WORKER) {
      return G.clients.matchAll({ includeUncontrolled: true, type: "window" }).then(function (cs) {
        cs.forEach(function (c) { c.postMessage({ type: kind }); });
      }).catch(function () {});
    }
    if (typeof G.onKeiSend === "function") { try { G.onKeiSend(kind); } catch (e) {} }
    return Promise.resolve();
  }

  /* ---------- the outbox ---------- */
  function list() {
    return db().then(function (d) {
      return new Promise(function (res, rej) {
        var out = [], tx = d.transaction("photos", "readonly");
        var rq = tx.objectStore("photos").openCursor(IDBKeyRange.bound("ob:", "ob:￿"));
        rq.onsuccess = function () { var c = rq.result; if (c) { out.push(c.value); c.continue(); } };
        tx.oncomplete = function () { out.sort(function (a, b) { return a.queued < b.queued ? -1 : 1; }); res(out); };
        tx.onerror = function () { rej(tx.error); };
      });
    });
  }
  function entry(id) { return get("ob:" + id); }
  function upd(id, patch) {
    return op("readwrite", function (st) {
      var rq = st.get("ob:" + id);
      rq.onsuccess = function () {
        var e = rq.result;
        if (!e) return;
        Object.keys(patch).forEach(function (k) { e[k] = patch[k]; });
        st.put(e, "ob:" + id);
      };
    }).then(function () { return notify("kei-outbox"); });
  }
  /* change an entry only while it is still the version a send started with; if the sheet was
     finished again or removed meanwhile, the send stops here ("superseded") and writes nothing */
  var SUPERSEDED = "superseded";
  function updRev(id, rev, patch) {
    var gone = false;
    return op("readwrite", function (st) {
      var rq = st.get("ob:" + id);
      rq.onsuccess = function () {
        var e = rq.result;
        if (!e || (e.rev || "") !== rev) { gone = true; return; }
        Object.keys(patch).forEach(function (k) { e[k] = patch[k]; });
        st.put(e, "ob:" + id);
      };
    }).then(function () {
      if (gone) throw new Error(SUPERSEDED);
      return notify("kei-outbox");
    });
  }
  function hexOf(buf) {
    return Array.prototype.map.call(new Uint8Array(buf), function (b) { return ("0" + b.toString(16)).slice(-2); }).join("");
  }
  /* a finished sheet: its file and its entry are written in ONE transaction. A sheet finished
     again after a correction keeps where its earlier copy went, so that copy is replaced. */
  function queueHtml(draftId, info, html) {
    var bytes = new TextEncoder().encode(html);
    return crypto.subtle.digest("SHA-256", bytes).then(function (h) {
      var blob = new Blob([bytes], { type: "text/html" });
      return op("readwrite", function (st) {
        var rq = st.get("ob:" + draftId);
        rq.onsuccess = function () {
          var old = rq.result || {};
          st.put(blob, "outbox:" + draftId);
          st.put({
            draftId: draftId, name: info.name, site: info.site, tech: info.tech || "", date: info.date || "",
            sha: hexOf(h), bytes: bytes.length, queued: new Date().toISOString(), state: "waiting", msg: "",
            siteDone: false, procDone: false,
            siteItemId: old.siteItemId || "", siteName: old.siteName || "", procStaged: !!old.procStaged,
            procStarted: !!(old.procStarted || old.procStaged), siteUp: null, procUp: null,
            rev: hexOf(h).slice(0, 12) + "-" + Date.now().toString(36)      /* this version of the sheet */
          }, "ob:" + draftId);
        };
      });
    }).then(function () { return notify("kei-outbox"); });
  }

  /* ---------- sign-in ---------- */
  function authGet() { return get("meta:auth").then(function (a) { return a || null; }); }
  function authPut(a) {
    return (a ? put("meta:auth", a) : del("meta:auth")).then(function () { return notify("kei-auth"); });
  }
  function postForm(url, data) {
    var body = Object.keys(data).map(function (k) { return encodeURIComponent(k) + "=" + encodeURIComponent(data[k]); }).join("&");
    return fetch(url, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body })
      .then(function (r) { return r.json(); });
  }
  function claimsOf(tok) {
    var out = { name: "", email: "" };
    try {
      var p = tok.id_token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
      var c = JSON.parse(decodeURIComponent(escape(atob(p + "===".slice((p.length + 3) % 4)))));
      out.name = c.name || ""; out.email = (c.preferred_username || c.upn || "").toLowerCase();
    } catch (e) {}
    return out;
  }
  /* a token response -> the stored sign-in. fresh = a new sign-in (restarts the 24 hours) */
  function keepTok(tok, fresh, old) {
    if (!tok || !tok.access_token) return Promise.resolve(null);
    var who = claimsOf(tok);
    old = old || {};
    var a = {
      access_token: tok.access_token,
      refresh_token: tok.refresh_token || old.refresh_token || "",
      expires_at: Date.now() + Number(tok.expires_in || 3600) * 1000,
      name: who.name || old.name || "", email: who.email || old.email || "",
      signedInAt: fresh ? Date.now() : (old.signedInAt || Date.now())
    };
    return authPut(a).then(function () { return a; });
  }
  function redeemCode(code, verifier, redirectUri) {
    return postForm(AUTHORITY + "/token", {
      client_id: CFG.clientId, grant_type: "authorization_code", code: code,
      redirect_uri: redirectUri, code_verifier: verifier, scope: SCOPES
    }).then(function (tok) {
      return keepTok(tok, true, null).then(function (a) {
        if (!a) throw new Error(tok.error_description || tok.error || "sign-in failed");
        return a;
      });
    });
  }
  var refreshing = null;
  function token() {
    return authGet().then(function (a) {
      if (!a) throw new Error(SIGNED_OUT);
      if (a.expires_at - 120000 > Date.now()) return a.access_token;
      if (refreshing) return refreshing;
      if (!a.refresh_token) return authPut(null).then(function () { throw new Error(SIGNED_OUT); });
      refreshing = postForm(AUTHORITY + "/token", {
        client_id: CFG.clientId, grant_type: "refresh_token", refresh_token: a.refresh_token, scope: SCOPES
      }).then(function (tok) {
        refreshing = null;
        if (!tok.access_token) return authPut(null).then(function () { throw new Error(SIGNED_OUT); });
        return keepTok(tok, false, a).then(function (n) { return n.access_token; });
      }, function (e) { refreshing = null; throw e; });   /* no signal is NOT a sign-out */
      return refreshing;
    });
  }

  /* ---------- Microsoft Graph ---------- */
  function gErr(status, msg) { var e = new Error(msg); e.status = status; return e; }
  function graph(method, path, body, ctype) {
    return token().then(function (t) {
      var h = { Authorization: "Bearer " + t };
      if (body != null) h["Content-Type"] = ctype || "application/json";
      return fetch((/^https?:/.test(path) ? path : CFG.graph + path), {
        method: method, headers: h,
        body: body == null ? undefined : (typeof body === "string" || body instanceof Blob ? body : JSON.stringify(body))
      });
    }).then(function (r) {
      if (r.status === 204) return {};
      return r.text().then(function (txt) {
        var j = {}; try { j = txt ? JSON.parse(txt) : {}; } catch (e) {}
        if (!r.ok) {
          var m = (j.error && j.error.message) || ("HTTP " + r.status);
          if (r.status === 401) return authPut(null).then(function () { throw new Error(SIGNED_OUT); });
          if (r.status === 403) m = "no permission to write there (" + m + ")";
          throw gErr(r.status, m);
        }
        return j;
      });
    });
  }
  function seg(x) { return encodeURIComponent(x); }
  function driveItem(id) { return "/drives/" + seg(CFG.driveId) + "/items/" + id; }
  function childPath(parentId, name) { return driveItem(parentId) + ":/" + seg(name) + ":"; }
  function itemOrNull(path) {
    return graph("GET", path + "?$select=id,name,size,folder,parentReference").catch(function (e) {
      if (e.status === 404) return null; throw e;
    });
  }
  /* Upload in 1.6 MB pieces. With rk ({id, field}) the upload session is kept in the outbox
     entry, so a send cut off by lost signal or a killed worker carries on from the last piece
     that arrived instead of starting over - the difference between a 16 MB sheet getting
     through on patchy signal or never. */
  function uploadBlob(sessionPath, behavior, blob, rk) {
    var size = blob.size, CH = 327680 * 5;
    if (!size) return Promise.reject(new Error("empty file"));
    function keep(v) { if (!rk) return Promise.resolve(); var p = {}; p[rk.field] = v; return rk.save(p); }
    function piece(url, at) {
      var end = Math.min(at + CH, size);
      return fetch(url, {
        method: "PUT", headers: { "Content-Range": "bytes " + at + "-" + (end - 1) + "/" + size },
        body: blob.slice(at, end)
      }).then(function (r) {
        if (r.status === 202) {
          return (rk ? keep({ url: url, path: sessionPath, size: size, at: end }) : Promise.resolve())
            .then(function () { return piece(url, end); });
        }
        return r.text().then(function (txt) {
          var j = {}; try { j = JSON.parse(txt); } catch (e) {}
          if (r.status === 200 || r.status === 201) return j;
          /* the session is gone (expired, or a piece it won't take): start over cleanly */
          if (r.status === 404 || r.status === 416) return fresh();
          throw gErr(r.status, (j.error && j.error.message) || ("upload failed, HTTP " + r.status));
        });
      });
    }
    function fresh() {
      return graph("POST", sessionPath + "/createUploadSession",
        { item: { "@microsoft.graph.conflictBehavior": behavior } }).then(function (sess) {
        var url = sess.uploadUrl;
        return (rk ? keep({ url: url, path: sessionPath, size: size, at: 0 }) : Promise.resolve())
          .then(function () { return piece(url, 0); });
      });
    }
    function resume(s) {
      return fetch(s.url).then(function (r) {
        if (!r.ok) return fresh();
        return r.json().then(function (j) {
          var nr = (j.nextExpectedRanges || [])[0], at = nr ? parseInt(nr.split("-")[0], 10) : NaN;
          return (at >= 0 && at < size) ? piece(s.url, at) : fresh();
        });
      });
    }
    var go = rk
      ? entry(rk.id).then(function (e) {
          var s = e && e[rk.field];
          return (s && s.url && s.path === sessionPath && s.size === size) ? resume(s) : fresh();
        })
      : fresh();
    return go.then(function (item) {
      return (rk ? keep(null) : Promise.resolve()).then(function () { return item; });
    });
  }
  /* The sheet's file name is decided ONCE, on its first try, and kept in its outbox entry:
     a name already used by another sheet gets _2, _3... Every later try reuses that name and
     may overwrite only this sheet's own unfinished copy. (Picking a fresh name on each retry
     is how Selous LB 2026-11-12 ended up as _5.) */
  function chooseName(parentId, e, save) {
    if (e.siteName) return Promise.resolve(e.siteName);
    function tryN(n) {
      var nm = n ? e.name.replace(/\.html$/i, "_" + (n + 1) + ".html") : e.name;
      return itemOrNull(childPath(parentId, nm)).then(function (it) {
        if (!it) return nm;
        if (n >= 8) throw new Error("too many files already named " + e.name);
        return tryN(n + 1);
      });
    }
    return tryN(0).then(function (nm) { return save({ siteName: nm }).then(function () { return nm; }); });
  }
  var sitesFolder = null;
  function sitesFolderId() {
    if (sitesFolder) return Promise.resolve(sitesFolder);
    return graph("GET", driveItem(CFG.staged) + "?$select=id,name,parentReference").then(function (it) {
      sitesFolder = it.parentReference && it.parentReference.id;
      if (!sitesFolder) throw new Error("could not find the Hydro Processor's Sites folder");
      return sitesFolder;
    });
  }
  function ensureFolder(parentId, name) {
    return itemOrNull(childPath(parentId, name)).then(function (it) {
      if (it) return it.id;
      return graph("POST", driveItem(parentId) + "/children",
        { name: name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
        .then(function (c) { return c.id; }, function (e) {
          if (e.status === 409) return itemOrNull(childPath(parentId, name)).then(function (x) { return x.id; });
          throw e;
        });
    });
  }
  function procId(site) { return PROC_ID[site] || site; }

  function sendToSite(e, blob, save) {
    var replace = e.siteItemId
      ? itemOrNull(driveItem(e.siteItemId)).then(function (it) {
          /* finished again after a correction: replace this sheet's own earlier upload */
          return it ? uploadBlob(driveItem(e.siteItemId), "replace", blob, { id: e.draftId, field: "siteUp", save: save }) : null;
        })
      : Promise.resolve(null);
    return replace.then(function (done) {
      if (done) return done;
      return itemOrNull(driveItem(CFG.monSites) + ":/" + seg(e.site) + "/" + seg("Field Sheets") + ":").then(function (f) {
        if (!f || !f.folder) throw new Error('there is no "Field Sheets" folder for ' + e.site + " in _KEI Monitoring Sites");
        return chooseName(f.id, e, save).then(function (nm) {
          return uploadBlob(childPath(f.id, nm), "replace", blob, { id: e.draftId, field: "siteUp", save: save });
        });
      });
    });
  }
  function sendToProcessor(e, blob, who, save) {
    var pid = procId(e.site);
    return sitesFolderId().then(function (sid) {
      return itemOrNull(childPath(sid, pid + ".keisite.json"));
    }).then(function (st) {
      if (!st) return "none";
      /* the Processor copy carries the same name as the OneDrive copy (a second visit that day is _2 in both) */
      var nm = e.siteName || e.name;
      return ensureFolder(CFG.staged, pid).then(function (fid) {
        return itemOrNull(childPath(fid, nm)).then(function (it) {
          if (it && it.size === blob.size) return it;
          if (it && !e.procStaged && !e.procStarted)
            throw new Error("another " + nm + " is still waiting in the Hydro Processor's approval queue - approve or set it aside there, then this one sends");
          return save({ procStarted: true }).then(function () {
            return uploadBlob(childPath(fid, nm), "replace", blob, { id: e.draftId, field: "procUp", save: save });
          });
        }).then(function () {
          var stage = {
            format: "kei-stage/1", station: pid, file: nm, kind: "fieldSheet", label: "Field sheet",
            why: "Sent from the phone app by " + ((who && (who.name || who.email)) || "the phone app"),
            sha256: e.sha, bytes: e.bytes, source: "phone:" + ((who && who.email) || ""),
            stagedUtc: new Date().toISOString().replace(/\.\d+Z$/, "Z")
          };
          /* written last: a stage is complete once its .stage.json exists */
          return graph("PUT", childPath(fid, nm + ".stage.json") + "/content", JSON.stringify(stage, null, 1), "application/json");
        }).then(function () { return "staged"; });
      });
    });
  }

  /* ---------- build 33: "a sheet is on its way" notice ----------
     A ~1 KB file in the Processor's queue folder, written BEFORE the big uploads. Best effort: nothing about
     it may ever stop or delay the sheet itself, so every failure here is swallowed. */
  var NOTICE_EVERY = 120000;            /* progress refresh at most every 2 minutes (G.KEI_NOTICE_EVERY: tests only) */
  function noticeName(e) { return e.draftId + ".waiting.json"; }
  function noticeFolder(e) {
    var pid = procId(e.site);
    return sitesFolderId().then(function (sid) {
      return itemOrNull(childPath(sid, pid + ".keisite.json"));
    }).then(function (st) { return st ? ensureFolder(CFG.staged, pid) : null; });
  }
  function putNotice(e, who, sent) {
    return noticeFolder(e).then(function (fid) {
      if (!fid) return false;
      var n = {
        format: "kei-waiting/1", station: procId(e.site), draftId: e.draftId, file: e.siteName || e.name,
        site: e.site, date: e.date || "", by: e.tech || "",
        sender: (who && (who.name || who.email)) || "", finishedUtc: e.queued || "",
        bytes: e.bytes || 0, totalBytes: (e.bytes || 0) * 2, sentBytes: Math.max(0, Math.round(sent || 0)),
        sha256: e.sha || "", updatedUtc: new Date().toISOString().replace(/\.\d+Z$/, "Z")
      };
      return graph("PUT", childPath(fid, noticeName(e)) + "/content", JSON.stringify(n, null, 1), "application/json")
        .then(function () { return true; });
    }).catch(function () { return false; });
  }
  function dropNotice(e) {
    var path = driveItem(CFG.staged) + ":/" + seg(procId(e.site)) + "/" + seg(noticeName(e)) + ":";
    return graph("DELETE", path).then(function () { return true; }, function (err) {
      return err && err.status === 404;              /* already gone counts as gone */
    });
  }

  function online() { return !G.navigator || G.navigator.onLine !== false; }
  function sendOne(id) {
    return entry(id).then(function (e) {
      if (!e || e.state === "sent") return null;
      var rev = e.rev || "";
      function save(patch) { return updRev(id, rev, patch); }
      /* build 33: the notice. noticeQ keeps its writes in order, so a late progress refresh can never land
         after the notice was taken away. */
      var noticeQ = Promise.resolve(), noticeAt = 0, siteAt = 0, procAt = 0, base = e.siteDone ? (e.bytes || 0) : 0, live = null;
      function notice(sent) {
        var l = live;                                 /* this sheet as it is now: never read later, when it may be gone */
        if (!l) return noticeQ;
        noticeAt = Date.now();
        noticeQ = noticeQ.then(function () { return putNotice(l, l.who, sent); }).then(function (ok) {
          return ok && !l.noticeUp ? (l.noticeUp = true, save({ noticeUp: true, noticeGone: false }).catch(function () {})) : null;
        }).catch(function () { return null; });       /* the queue itself never fails */
        return noticeQ;
      }
      function saveUp(patch) {
        return save(patch).then(function () {
          if (patch.siteUp && patch.siteUp.at != null) siteAt = patch.siteUp.at;
          if (patch.procUp && patch.procUp.at != null) procAt = patch.procUp.at;
          if (live && Date.now() - noticeAt >= (G.KEI_NOTICE_EVERY || NOTICE_EVERY)) notice(base + siteAt + procAt);   /* not waited for */
        });
      }
      return save({ state: "sending", msg: "" }).then(function () {
        return Promise.all([get("outbox:" + id), authGet()]);
      }).then(function (got) {
        var blob = got[0], who = got[1];
        if (!blob) throw new Error("the finished file is missing from this phone - open the sheet and press Finish again");
        return entry(id).then(function (cur) {
          if (!cur || (cur.rev || "") !== rev) throw new Error(SUPERSEDED);
          live = cur; live.who = who;
          return (cur.procDone ? null : notice(base)).then(function () {
            return cur.siteDone ? null : sendToSite(cur, blob, saveUp).then(function (it) {
              live.siteName = it.name || cur.name; base = cur.bytes || 0; siteAt = 0;
              return save({ siteDone: true, siteItemId: it.id, siteName: it.name || cur.name });
            });
          });
        }).then(function () {
          return entry(id).then(function (c2) {
            if (!c2 || (c2.rev || "") !== rev) throw new Error(SUPERSEDED);
            return c2.procDone ? null : sendToProcessor(c2, blob, who, saveUp).then(function (r) {
              return save({ procDone: true, proc: r, procStaged: r === "staged" || c2.procStaged });
            });
          });
        });
      }).then(function () {
        var l = live; live = null;                    /* no more refreshes */
        return noticeQ.then(function () {
          return entry(id).then(function (c3) {
            return (c3 && c3.noticeUp && !c3.noticeGone) || (l && l.noticeUp) ? dropNotice(l).then(function (gone) {
              return gone ? save({ noticeGone: true }) : null;
            }) : null;
          });
        }).catch(function () { return null; }).then(function () {
          return save({ state: "sent", msg: "", sentAt: new Date().toISOString() });
        });
      }).catch(function (err) {
        var m = err && err.message ? err.message : String(err);
        if (m === SUPERSEDED) return null;          /* finished again or removed: the next pass sends the new one */
        var p = m === SIGNED_OUT ? { state: "signin", msg: "" }
          : (!online() || err instanceof TypeError) ? { state: "waiting", msg: "" }
          : { state: "error", msg: m };
        return save(p).catch(function () { return null; });
      });
    });
  }
  /* Send everything still to go, one at a time. Calls while a run is going share that run, and
     the run keeps taking passes until nothing NEW is waiting - so a sheet finished while another
     is uploading goes out as soon as that one is done, not at the next trigger.
     Resolves to what is left: {left, waiting, signin, errors}. */
  var running = null;
  function pumpAll() {
    if (running) return running;
    var tried = {};
    function key(x) { return x.draftId + "|" + (x.rev || ""); }
    function pass(n) {
      return list().then(function (l) {
        var todo = l.filter(function (x) { return x.state !== "sent" && !tried[key(x)]; });
        if (!todo.length || n >= 25) return null;
        if (!online()) {
          return Promise.all(todo.map(function (x) {
            return x.state === "waiting" ? null : upd(x.draftId, { state: "waiting", msg: "" });
          }));
        }
        return authGet().then(function (a) {
          if (!a) {
            return Promise.all(todo.map(function (x) {
              return x.state === "signin" ? null : upd(x.draftId, { state: "signin", msg: "" });
            }));
          }
          return todo.reduce(function (chain, x) {
            return chain.then(function () { tried[key(x)] = 1; return sendOne(x.draftId); });
          }, Promise.resolve()).then(function () { return pass(n + 1); });
        });
      });
    }
    running = pass(0).then(list).then(function (l) {
      var left = l.filter(function (x) { return x.state !== "sent"; });
      running = null;
      return {
        left: left.length,
        waiting: left.filter(function (x) { return x.state === "waiting" || x.state === "sending"; }).length,
        signin: left.filter(function (x) { return x.state === "signin"; }).length,
        errors: left.filter(function (x) { return x.state === "error"; }).length
      };
    }, function (e) { running = null; throw e; });
    return running;
  }

  /* take a sheet out of the outbox for good: its entry and its file, together */
  function remove(id) {
    return entry(id).then(function (e) {
      return op("readwrite", function (st) { st.delete("ob:" + id); st.delete("outbox:" + id); })
        .then(function () { return notify("kei-outbox"); })
        .then(function () {                          /* build 33: its notice, best effort, not waited on by the page */
          if (e && e.noticeUp && !e.noticeGone) dropNotice(e).catch(function () {});
        });
    });
  }

  /* Has the office dealt with this sheet in the Hydro Processor? Reads only. An upload waits in
     _Staged\<station>\ until someone decides; the Processor then files it in
     _Staged\_Done\<station>\ as "<time> <file>" and writes the decision into its .stage.json.
     Only a record of THIS file (same sha256) counts. Result is kept in the entry:
       procState  "queued" | "accepted" | "already in the record" | "set aside" | "missing"  */
  function listAll(path) {
    var out = [];
    function page(p, n) {
      return graph("GET", p).then(function (r) {
        out = out.concat(r.value || []);
        var next = r["@odata.nextLink"];
        return (next && n < 10) ? page(next, n + 1) : out;
      });
    }
    return page(path, 0);
  }
  function checkProcessor(id) {
    return entry(id).then(function (e) {
      if (e && e.state === "sent" && e.noticeUp && !e.noticeGone) {      /* build 33: a notice left behind */
        dropNotice(e).then(function (gone) { return gone ? upd(id, { noticeGone: true }) : null; }).catch(function () {});
      }
      if (!e || e.state !== "sent" || e.proc !== "staged") return null;
      var pid = procId(e.site), nm = e.siteName || e.name, base = driveItem(CFG.staged);
      var tail = " " + nm + ".stage.json";
      return itemOrNull(base + ":/" + seg(pid) + "/" + seg(nm + ".stage.json") + ":").then(function (waiting) {
        if (waiting) return { procState: "queued", procDecided: "" };
        return listAll(base + ":/_Done/" + seg(pid) + ":/children?$top=200").catch(function (err) {
          if (err.status === 404) return [];
          throw err;
        }).then(function (kids) {
          var hits = kids.filter(function (k) {
            return k.name && k.name.length > tail.length && k.name.slice(-tail.length) === tail;
          }).sort(function (a, b) { return a.name < b.name ? 1 : -1; });       /* newest first */
          function read(i) {
            if (i >= hits.length) return { procState: "missing", procDecided: "" };
            var url = hits[i]["@microsoft.graph.downloadUrl"];
            if (!url) return read(i + 1);
            return fetch(url).then(function (r) { return r.json(); }).then(function (j) {
              if (j && j.sha256 === e.sha) return { procState: String(j.decision || "filed"), procDecided: j.decidedUtc || "" };
              return read(i + 1);
            });
          }
          return read(0);
        });
      }).then(function (res) {
        res.procChecked = new Date().toISOString();
        return upd(id, res).then(function () { return res; });
      });
    });
  }

  G.KEISEND = {
    CFG: CFG, AUTHORITY: AUTHORITY, SCOPES: SCOPES, SIGNED_OUT: SIGNED_OUT,
    list: list, entry: entry, upd: upd, queueHtml: queueHtml, pumpAll: pumpAll,
    authGet: authGet, authPut: authPut, redeemCode: redeemCode, token: token,
    graph: graph, itemOrNull: itemOrNull, childPath: childPath, driveItem: driveItem, seg: seg,
    sitesFolderId: sitesFolderId, procId: procId, getFile: function (id) { return get("outbox:" + id); },
    resetFolderCache: function () { sitesFolder = null; },
    remove: remove, checkProcessor: checkProcessor,
    putNotice: putNotice, dropNotice: dropNotice      /* build 33 */
  };
})(self);

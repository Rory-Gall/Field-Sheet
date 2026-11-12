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
            procStarted: !!(old.procStarted || old.procStaged), siteUp: null, procUp: null
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
      return fetch(CFG.graph + path, {
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
    function keep(v) { if (!rk) return Promise.resolve(); var p = {}; p[rk.field] = v; return upd(rk.id, p); }
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
  function chooseName(parentId, e) {
    if (e.siteName) return Promise.resolve(e.siteName);
    function tryN(n) {
      var nm = n ? e.name.replace(/\.html$/i, "_" + (n + 1) + ".html") : e.name;
      return itemOrNull(childPath(parentId, nm)).then(function (it) {
        if (!it) return nm;
        if (n >= 8) throw new Error("too many files already named " + e.name);
        return tryN(n + 1);
      });
    }
    return tryN(0).then(function (nm) { return upd(e.draftId, { siteName: nm }).then(function () { return nm; }); });
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

  function sendToSite(e, blob) {
    var replace = e.siteItemId
      ? itemOrNull(driveItem(e.siteItemId)).then(function (it) {
          /* finished again after a correction: replace this sheet's own earlier upload */
          return it ? uploadBlob(driveItem(e.siteItemId), "replace", blob, { id: e.draftId, field: "siteUp" }) : null;
        })
      : Promise.resolve(null);
    return replace.then(function (done) {
      if (done) return done;
      return itemOrNull(driveItem(CFG.monSites) + ":/" + seg(e.site) + "/" + seg("Field Sheets") + ":").then(function (f) {
        if (!f || !f.folder) throw new Error('there is no "Field Sheets" folder for ' + e.site + " in _KEI Monitoring Sites");
        return chooseName(f.id, e).then(function (nm) {
          return uploadBlob(childPath(f.id, nm), "replace", blob, { id: e.draftId, field: "siteUp" });
        });
      });
    });
  }
  function sendToProcessor(e, blob, who) {
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
          return upd(e.draftId, { procStarted: true }).then(function () {
            return uploadBlob(childPath(fid, nm), "replace", blob, { id: e.draftId, field: "procUp" });
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

  function online() { return !G.navigator || G.navigator.onLine !== false; }
  function sendOne(id) {
    return entry(id).then(function (e) {
      if (!e || e.state === "sent") return null;
      return upd(id, { state: "sending", msg: "" }).then(function () {
        return Promise.all([get("outbox:" + id), authGet()]);
      }).then(function (got) {
        var blob = got[0], who = got[1];
        if (!blob) throw new Error("the finished file is missing from this phone - open the sheet and press Finish again");
        return entry(id).then(function (cur) {
          return cur.siteDone ? null : sendToSite(cur, blob).then(function (it) {
            return upd(id, { siteDone: true, siteItemId: it.id, siteName: it.name || cur.name });
          });
        }).then(function () {
          return entry(id).then(function (c2) {
            return c2.procDone ? null : sendToProcessor(c2, blob, who).then(function (r) {
              return upd(id, { procDone: true, proc: r, procStaged: r === "staged" || c2.procStaged });
            });
          });
        });
      }).then(function () {
        return upd(id, { state: "sent", msg: "", sentAt: new Date().toISOString() });
      }).catch(function (err) {
        var m = err && err.message ? err.message : String(err);
        if (m === SIGNED_OUT) return upd(id, { state: "signin", msg: "" });
        if (!online() || err instanceof TypeError) return upd(id, { state: "waiting", msg: "" });
        return upd(id, { state: "error", msg: m });
      });
    });
  }
  /* send everything still to go, one at a time. Calls while a run is going share that run.
     Resolves to what is left: {left, waiting, signin, errors}. */
  var running = null;
  function pumpAll() {
    if (running) return running;
    running = list().then(function (l) {
      var todo = l.filter(function (x) { return x.state !== "sent"; });
      if (!todo.length) return null;
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
          return chain.then(function () { return sendOne(x.draftId); });
        }, Promise.resolve());
      });
    }).then(list).then(function (l) {
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

  G.KEISEND = {
    CFG: CFG, AUTHORITY: AUTHORITY, SCOPES: SCOPES, SIGNED_OUT: SIGNED_OUT,
    list: list, entry: entry, upd: upd, queueHtml: queueHtml, pumpAll: pumpAll,
    authGet: authGet, authPut: authPut, redeemCode: redeemCode, token: token,
    graph: graph, itemOrNull: itemOrNull, childPath: childPath, driveItem: driveItem, seg: seg,
    sitesFolderId: sitesFolderId, procId: procId, getFile: function (id) { return get("outbox:" + id); },
    resetFolderCache: function () { sitesFolder = null; }
  };
})(self);

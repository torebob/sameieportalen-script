document.addEventListener('DOMContentLoaded', () => {
  const get = (id) => document.getElementById(id);

  const dom = {
    meetingCard: get('meetingCard'),
    agendaCard: get('agendaCard'),
    statusBadge: get('statusBadge'),
    btnUnlock: get('btnUnlock'),
    liveBadge: get('liveBadge'),
    moteId: get('mote-id'),
    type: get('type'),
    dato: get('dato'),
    start: get('start'),
    slutt: get('slutt'),
    sted: get('sted'),
    tittel: get('tittel'),
    conflictBanner: get('conflictBanner'),
    btnReload: get('btnReload'),
    btnSaveMeeting: get('btnSaveMeeting'),
    btnNewMeeting: get('btnNewMeeting'),
    saveStatus: get('saveStatus'),
    scopePlanned: get('scopePlanned'),
    scopePast: get('scopePast'),
    planned: get('planned'),
    btnRefresh: get('btnRefresh'),
    agendaCard: get('agendaCard'),
    btnNewSak: get('btnNewSak'),
    moveWrap: get('moveWrap'),
    moveTo: get('moveTo'),
    btnMoveSak: get('btnMoveSak'),
    saksnr: get('saksnr'),
    sakTittel: get('sakTittel'),
    innspillFrom: get('innspillFrom'),
    innspillText: get('innspillText'),
    btnNyttInnspill: get('btnNyttInnspill'),
    btnLagreSak: get('btnLagreSak'),
    btnSlettSak: get('btnSlettSak'),
    innspillList: get('innspillList'),
    forslagText: get('forslagText'),
    vedtakText: get('vedtakText'),
  };

  const LOCKED_STATUSES = ['Protokoll under godkjenning', 'Protokoll godkjent', 'Arkivert'];

  let state = {
    user: null,
    moteId: null,
    sakId: null,
    lastTick: null,
    pollTimer: null,
    pollMs: 5000,
    meetingStatus: null,
    canUnlock: false,
  };

  const setStatus = (msg) => {
    dom.saveStatus.textContent = msg || '';
  };

  const fmtDate = (d) => {
    try {
      return new Date(d).toLocaleDateString('no-NO');
    } catch (_) {
      return '';
    }
  };

  const toIsoDate = (val) => {
    const s = String(val || '').trim();
    if (!s) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) {
      return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    }
    return s;
  };

  const isLocked = (st) => !!st && LOCKED_STATUSES.includes(String(st).trim());

  const applyLockUI = (locked) => {
    [dom.agendaCard, dom.meetingCard].forEach((root) => {
      if (!root) return;
      root.classList.toggle('locked', !!locked);
      root.querySelectorAll('input,textarea,select').forEach((el) => {
        if (el.id === 'saksnr') return;
        el.disabled = !!locked;
      });
      root.querySelectorAll('button').forEach((btn) => {
        const id = btn.id || '';
        const whitelist = ['btnUnlock', 'btnReload', 'btnRefresh'];
        btn.disabled = !!locked && !whitelist.includes(id);
      });
    });
    dom.moveWrap.style.display = (state.sakId && !locked) ? 'inline-flex' : 'none';
  };

  const updateStatusBadge = () => {
    const b = dom.statusBadge;
    if (!state.meetingStatus) {
      b.style.display = 'none';
      return;
    }
    b.textContent = state.meetingStatus;
    b.className = 'pill' + (isLocked(state.meetingStatus) ? ' lock' : '');
    b.style.display = 'inline-block';
    dom.btnUnlock.style.display = (isLocked(state.meetingStatus) && state.canUnlock) ? 'inline-flex' : 'none';
  };

  const server = {
    bootstrap: (onSuccess) => google.script.run.withSuccessHandler(onSuccess).uiBootstrap(),
    listMeetings: (scope, onSuccess) => google.script.run.withSuccessHandler(onSuccess).listMeetings_({ scope }),
    upsertMeeting: (form, onSuccess, onFailure) => google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFailure).upsertMeeting(form),
    hasPermission: (permission, onSuccess) => google.script.run.withSuccessHandler(onSuccess).hasPermission(permission),
    listAgenda: (moteId, onSuccess) => google.script.run.withSuccessHandler(onSuccess).listAgenda(moteId),
    listInnspill: (sakId, onSuccess) => google.script.run.withSuccessHandler(onSuccess).listInnspill(sakId, null),
    unlockMeeting: (moteId, onSuccess, onFailure) => google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFailure).unlockMeeting(moteId),
    newAgendaItem: (moteId, onSuccess) => google.script.run.withSuccessHandler(onSuccess).newAgendaItem(moteId),
    saveAgenda: (payload, onSuccess, onFailure) => google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFailure).saveAgenda(payload),
    appendInnspill: (sakId, text, onSuccess, onFailure) => google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFailure).appendInnspill(sakId, text),
    moveAgendaToMeeting: (sakId, toMoteId, onSuccess, onFailure) => google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFailure).moveAgendaToMeeting(sakId, toMoteId),
    rtServerNow: (onSuccess) => google.script.run.withSuccessHandler(onSuccess).rtServerNow(),
    rtGetChanges: (moteId, lastTick, onSuccess) => google.script.run.withSuccessHandler(onSuccess).rtGetChanges(moteId, lastTick),
  };

  const refreshMeetings = () => {
    dom.planned.innerHTML = '<li class="hint">Laster…</li>';
    const scope = dom.scopePast.classList.contains('active') ? 'past' : 'planned';
    server.listMeetings(scope, drawMeetings);
  };

  const drawMeetings = (list) => {
    const ul = dom.planned;
    ul.innerHTML = '';
    if (!list || !list.length) {
      ul.innerHTML = '<li class="hint">Ingen møter i listen.</li>';
      return;
    }
    list.forEach((m) => {
      const li = document.createElement('li');
      li.innerHTML = `<strong>${escapeHtml(m.tittel || m.type)}</strong><br>${escapeHtml(m.type)} • ${m.dato ? fmtDate(m.dato) : ''} • ${escapeHtml(m.start || '')} • ${escapeHtml(m.sted || '')}`;
      li.style.cursor = 'pointer';
      li.addEventListener('click', () => loadMeeting(m));
      ul.appendChild(li);
    });
  };

  const loadMeeting = (m) => {
    state.moteId = m.id;
    state.meetingStatus = m.status || 'Planlagt';
    dom.moteId.value = state.moteId;
    dom.type.value = m.type || 'Styremøte';
    dom.dato.value = m.dato ? new Date(m.dato).toISOString().slice(0, 10) : '';
    dom.start.value = m.start || '';
    dom.slutt.value = m.slutt || '';
    dom.sted.value = m.sted || '';
    dom.tittel.value = m.tittel || '';
    setStatus('Møte lastet.');
    dom.conflictBanner.style.display = 'none';
    dom.liveBadge.style.display = 'none';
    dom.innspillList.innerHTML = '<div class="hint">Ingen innspill ennå.</div>';
    updateStatusBadge();
    applyLockUI(isLocked(state.meetingStatus));

    if (google.script.run.hasPermission) {
      server.hasPermission('MEETING_UNLOCK', (can) => {
        state.canUnlock = !!can;
        updateStatusBadge();
      });
    }

    server.listAgenda(state.moteId, (saker) => {
      if (!saker || !saker.length) {
        clearSak();
        startPolling();
        return;
      }
      const s = saker[0];
      state.sakId = s.sakId;
      dom.saksnr.value = s.saksnr || '';
      dom.sakTittel.value = s.tittel || '';
      dom.innspillText.value = '';
      dom.forslagText.value = s.forslag || '';
      dom.vedtakText.value = s.vedtak || '';
      startPolling();
      if (state.sakId) {
        if (google.script.run.listInnspill) {
          server.listInnspill(state.sakId, drawInnspill);
        }
        loadMoveTargets();
      }
    });
  };

  const drawInnspill = (items) => {
    const box = dom.innspillList;
    if (!items || !items.length) {
      if (box.children.length === 0) box.innerHTML = '<div class="hint">Ingen innspill ennå.</div>';
      return;
    }
    if (box.children.length === 1 && box.querySelector('.hint')) box.innerHTML = '';
    items.forEach((x) => {
      const el = document.createElement('div');
      el.className = 'innspill-item';
      const ts = x.ts ? new Date(x.ts).toLocaleString('no-NO') : '';
      el.innerHTML = `<div class="hint">${ts} • ${x.from || 'ukjent'}</div><div>${escapeHtml(x.text || '')}</div>`;
      box.appendChild(el);
    });
    box.scrollTop = box.scrollHeight;
  };

  const clearMeeting = () => {
    stopPolling();
    state.moteId = null;
    state.meetingStatus = null;
    state.canUnlock = false;
    dom.moteId.value = '';
    dom.type.value = 'Styremøte';
    dom.dato.value = '';
    dom.start.value = '18:00';
    dom.slutt.value = '20:00';
    dom.sted.value = '';
    dom.tittel.value = '';
    setStatus('');
    clearSak();
    updateStatusBadge();
    applyLockUI(false);
    dom.innspillList.innerHTML = '<div class="hint">Ingen innspill ennå.</div>';
  };

  const clearSak = () => {
    state.sakId = null;
    dom.saksnr.value = '';
    dom.sakTittel.value = '';
    dom.innspillText.value = '';
    dom.forslagText.value = '';
    dom.vedtakText.value = '';
    dom.moveWrap.style.display = 'none';
  };

  const saveMeeting = () => {
    const form = {
      moteId: dom.moteId.value || state.moteId || '',
      type: dom.type.value,
      datoISO: toIsoDate(dom.dato.value),
      start: dom.start.value,
      slutt: dom.slutt.value,
      sted: dom.sted.value,
      tittel: dom.tittel.value,
      agenda: '',
    };
    if (!form.tittel.trim()) {
      setStatus('Møtetittel er påkrevd.');
      return;
    }
    if (!form.datoISO) {
      setStatus('Velg dato.');
      return;
    }

    setStatus('Lagrer …');
    server.upsertMeeting(
      form,
      (res) => {
        if (res && res.ok) {
          state.moteId = res.id || state.moteId;
          dom.moteId.value = state.moteId || '';
          setStatus(res.message || 'Lagret.');
          refreshMeetings();
        } else {
          setStatus(res?.message || 'Kunne ikke lagre.');
        }
      },
      (err) => setStatus(err?.message || 'En feil oppstod.')
    );
  };

  const unlockMeeting = () => {
    if (!state.moteId) return;
    setStatus('Låser opp …');
    if (!google.script.run.unlockMeeting) {
      setStatus('Server mangler funksjonen unlockMeeting(moteId).');
      return;
    }
    server.unlockMeeting(
      state.moteId,
      (res) => {
        if (res && res.ok) {
          state.meetingStatus = res.status || 'Planlagt';
          updateStatusBadge();
          applyLockUI(false);
          setStatus('Møtet er låst opp.');
        } else {
          setStatus(res?.message || 'Kunne ikke låse opp.');
        }
      },
      (e) => setStatus(e?.message || 'Feil ved ulåsing.')
    );
  };

  const newSak = () => {
    if (!state.moteId) {
      setStatus('Lagre eller velg et møte først.');
      return;
    }
    setStatus('Oppretter ny sak...');
    server.newAgendaItem(state.moteId, (res) => {
      if (!res || !res.ok) {
        setStatus(res?.message || 'Kunne ikke opprette sak.');
        return;
      }
      state.sakId = res.sakId;
      dom.saksnr.value = res.saksnr;
      dom.sakTittel.value = '';
      dom.innspillText.value = '';
      dom.forslagText.value = '';
      dom.vedtakText.value = '';
      setStatus('Ny sak opprettet. Du kan nå skrive inn tittel og detaljer.');
      loadMoveTargets();
    });
  };

  const saveSak = () => {
    if (!state.sakId) {
      setStatus('Opprett eller velg en sak først.');
      return;
    }
    setStatus('Lagrer sak...');
    const payload = {
      sakId: state.sakId,
      tittel: dom.sakTittel.value,
      forslag: dom.forslagText.value,
      vedtak: dom.vedtakText.value,
    };
    server.saveAgenda(
      payload,
      (_) => setStatus('Sak lagret.'),
      (err) => setStatus(err?.message || 'Feil ved lagring.')
    );
  };

  const nyttInnspill = () => {
    if (!state.sakId) {
      setStatus('Opprett eller velg en sak først.');
      return;
    }
    const txt = (dom.innspillText.value || '').trim();
    if (!txt) {
      setStatus('Skriv et innspill først.');
      return;
    }
    setStatus('Sender innspill...');
    server.appendInnspill(
      state.sakId,
      txt,
      (_) => {
        dom.innspillText.value = '';
        setStatus('Innspill lagret.');
      },
      (err) => setStatus(err?.message || 'Feil ved lagring.')
    );
  };

  const loadMoveTargets = () => {
    if (!state.moteId) return;
    server.listMeetings('planned', (list) => {
      const sel = dom.moveTo;
      const wrap = dom.moveWrap;
      sel.innerHTML = '';
      const options = (list || []).filter((m) => m.id !== state.moteId);
      if (!state.sakId || options.length === 0) {
        wrap.style.display = 'none';
        return;
      }
      options.forEach((m) => {
        const opt = document.createElement('option');
        const d = m.dato ? new Date(m.dato).toLocaleDateString('no-NO') : '';
        opt.value = m.id;
        opt.textContent = `${m.tittel || m.type || 'Møte'} • ${d} ${m.start || ''}`;
        sel.appendChild(opt);
      });
      wrap.style.display = isLocked(state.meetingStatus) ? 'none' : 'inline-flex';
    });
  };

  const moveSak = () => {
    if (!state.sakId) {
      setStatus('Velg en sak først.');
      return;
    }
    const toId = dom.moveTo.value;
    if (!toId) {
      setStatus('Velg mål-møte.');
      return;
    }
    if (!google.script.run.moveAgendaToMeeting) {
      setStatus('Server mangler funksjonen moveAgendaToMeeting(sakId, toMoteId).');
      return;
    }
    setStatus('Flytter sak …');
    server.moveAgendaToMeeting(
      state.sakId,
      toId,
      (res) => {
        if (res && res.ok) {
          setStatus('Sak flyttet.');
          clearSak();
          dom.innspillList.innerHTML = '<div class="hint">Ingen innspill ennå.</div>';
        } else {
          setStatus(res?.message || 'Kunne ikke flytte sak.');
        }
      },
      (e) => setStatus(e?.message || 'Feil ved flytting.')
    );
  };

  const startPolling = () => {
    stopPolling();
    state.lastTick = null;
    if (!google.script.run.rtServerNow) return;
    server.rtServerNow((x) => {
      state.lastTick = x?.now || new Date().toISOString();
      state.pollMs = 5000;
      state.pollTimer = setInterval(doPoll, state.pollMs);
    });
  };

  const stopPolling = () => {
    if (state.pollTimer) {
      clearInterval(state.pollTimer);
      state.pollTimer = null;
    }
  };

  const adjustPolling = (hasActivity) => {
    state.pollMs = hasActivity ? 5000 : Math.min(30000, state.pollMs + 5000);
    if (state.pollTimer) {
      clearInterval(state.pollTimer);
      state.pollTimer = setInterval(doPoll, state.pollMs);
    }
  };

  const doPoll = () => {
    if (!state.moteId || !google.script.run.rtGetChanges) return;
    server.rtGetChanges(state.moteId, state.lastTick, (ch) => {
      if (!ch) return;
      let activity = false;
      if (ch.meetingUpdated) {
        dom.conflictBanner.style.display = 'block';
        dom.liveBadge.style.display = 'inline-block';
        activity = true;
        if (ch.meetingStatus) {
          state.meetingStatus = ch.meetingStatus;
          updateStatusBadge();
          applyLockUI(isLocked(state.meetingStatus));
        }
      }
      if (ch.updatedSaker && ch.updatedSaker.length) {
        const hit = ch.updatedSaker.find((s) => s.sakId === state.sakId);
        if (hit) {
          const ae = document.activeElement || {};
          const typing = (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA');
          if (!typing) {
            dom.sakTittel.value = hit.tittel ?? dom.sakTittel.value;
            dom.forslagText.value = hit.forslag ?? dom.forslagText.value;
            dom.vedtakText.value = hit.vedtak ?? dom.vedtakText.value;
          } else {
            dom.conflictBanner.style.display = 'block';
          }
          dom.liveBadge.style.display = 'inline-block';
          activity = true;
        }
      }
      if (ch.newInnspill && ch.newInnspill.length) {
        if (state.sakId) {
          const relevant = ch.newInnspill.filter((x) => x.sakId === state.sakId);
          if (relevant.length) {
            appendInnspillToList(relevant);
            dom.liveBadge.style.display = 'inline-block';
            activity = true;
          }
        }
      }
      state.lastTick = ch.serverNow || new Date().toISOString();
      adjustPolling(activity);
    });
  };

  const refreshCurrentMeeting = () => {
    server.listMeetings(
      dom.scopePast.classList.contains('active') ? 'past' : 'planned',
      (list) => {
        const m = (list || []).find((x) => x.id === state.moteId);
        if (m) {
          loadMeeting(m);
        }
        dom.conflictBanner.style.display = 'none';
        dom.liveBadge.style.display = 'none';
      }
    );
  };

  const appendInnspillToList = (items) => {
    const box = dom.innspillList;
    if (box.children.length === 1 && box.querySelector('.hint')) box.innerHTML = '';
    items.forEach((x) => {
      const el = document.createElement('div');
      el.className = 'innspill-item';
      const ts = x.ts ? new Date(x.ts).toLocaleString('no-NO') : '';
      el.innerHTML = `<div class="hint">${ts} • ${x.from || 'ukjent'}</div><div>${escapeHtml(x.text || '')}</div>`;
      box.appendChild(el);
    });
    box.scrollTop = box.scrollHeight;
  };

  // Event Listeners
  dom.btnSaveMeeting.addEventListener('click', saveMeeting);
  dom.btnNewMeeting.addEventListener('click', clearMeeting);
  dom.btnRefresh.addEventListener('click', refreshMeetings);
  dom.btnReload.addEventListener('click', () => {
    if (state.moteId) refreshCurrentMeeting();
  });
  dom.scopePlanned.addEventListener('click', () => {
    dom.scopePlanned.classList.add('active');
    dom.scopePast.classList.remove('active');
    refreshMeetings();
  });
  dom.scopePast.addEventListener('click', () => {
    dom.scopePast.classList.add('active');
    dom.scopePlanned.classList.remove('active');
    refreshMeetings();
  });
  dom.btnUnlock.addEventListener('click', unlockMeeting);
  dom.btnNewSak.addEventListener('click', newSak);
  dom.btnLagreSak.addEventListener('click', saveSak);
  dom.btnNyttInnspill.addEventListener('click', nyttInnspill);
  dom.btnMoveSak.addEventListener('click', moveSak);

  // Initial load
  server.bootstrap((res) => {
    state.user = res?.user || null;
    dom.innspillFrom.textContent = 'Fra: ' + (state.user?.email || '—');
    refreshMeetings();
  });
});
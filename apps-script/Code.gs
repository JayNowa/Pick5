var FIREBASE_PROJECT_ID = 'pick5-ee939';
var FIREBASE_API_KEY    = 'AIzaSyC6wED6bBvMNGiQFFEGHdQGSP7iKGSIdvk';
var APP_URL             = 'https://pick5game.win';
var ADMIN_EMAILS        = ['jenowakoski@gmail.com', 'wcgriffiths14@gmail.com'];
var SHEET_ID            = '1wCbkfJMaFQoV3gYwFbxrijqLsxpLUXdlK4YM52ijrPE';

// ─── ENTRY POINTS ────────────────────────────────────────────────────────────

function doGet(e) { return doPost(e); }

function doPost(e) {
  console.log('doPost called - v12');
  try {
    var raw = (e.postData && e.postData.contents)
           || (e.parameter && e.parameter.data)
           || '{}';

    var data = JSON.parse(raw);

    if (data.type === 'weeklyRecap')   { sendWeeklyRecapEmail(data);   return ContentService.createTextOutput('Recap sent'); }
    if (data.type === 'picksReminder') { sendPicksReminderEmail(data); return ContentService.createTextOutput('Reminder sent'); }

    logPicksToSheet(data);
    sendPicksConfirmation(data);
    return ContentService.createTextOutput('OK');

  } catch (err) {
    Logger.log('doPost error: ' + err);
    MailApp.sendEmail('jenowakoski@gmail.com', 'doPost ERROR', err.toString());
    return ContentService.createTextOutput('Error: ' + err);
  }
}

// ─── SHEET LOGGING ───────────────────────────────────────────────────────────

function logPicksToSheet(data) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Picks') || ss.insertSheet('Picks');
    sheet.appendRow([
      new Date(), data.teamName || data.name || '', data.email || '', data.week || '',
      data.pick1 || '', data.pick2 || '', data.pick3 || '', data.pick4 || '', data.pick5 || '',
      data.spread1 || '', data.spread2 || '', data.spread3 || '', data.spread4 || '', data.spread5 || '',
      data.tiebreaker !== undefined ? data.tiebreaker : '',
      data.tiebreakerGame || ''
    ]);
  } catch (e) {
    Logger.log('logPicksToSheet error: ' + e);
  }
}

// ─── PICKS CONFIRMATION ──────────────────────────────────────────────────────

function sendPicksConfirmation(data) {
  try {
    const slugs = {
      'ARI':'ari','ATL':'atl','BAL':'bal','BUF':'buf','CAR':'car','CHI':'chi',
      'CIN':'cin','CLE':'cle','DAL':'dal','DEN':'den','DET':'det','GB':'gb',
      'HOU':'hou','IND':'ind','JAX':'jax','KC':'kc','LAC':'lac','LAR':'lar',
      'LV':'lv','MIA':'mia','MIN':'min','NE':'ne','NO':'no','NYG':'nyg',
      'NYJ':'nyj','PHI':'phi','PIT':'pit','SF':'sf','SEA':'sea','TB':'tb',
      'TEN':'ten','WSH':'wsh'
    };

    function logoUrl(abbr) {
      const slug = abbr ? slugs[abbr.trim().toUpperCase()] : null;
      return slug
        ? 'https://a.espncdn.com/combiner/i?img=/i/teamlogos/nfl/500/' + slug + '.png&w=32&h=32&transparent=true'
        : '';
    }

    const picks = [
      { team: data.pick1, spread: data.spread1 },
      { team: data.pick2, spread: data.spread2 },
      { team: data.pick3, spread: data.spread3 },
      { team: data.pick4, spread: data.spread4 },
      { team: data.pick5, spread: data.spread5 }
    ].filter(p => p.team && p.team.trim() !== '');

    const picksList = picks.map((p, i) => {
      const logo = logoUrl(p.team);
      const logoHtml = logo
        ? `<img src="${logo}" width="28" height="28" style="vertical-align:middle;margin-right:8px;">`
        : '';
      return `<tr>
        <td style="padding:10px 16px;border-bottom:1px solid #2a3040;color:#6b7588;font-family:monospace;font-size:14px;">${i + 1}</td>
        <td style="padding:10px 16px;border-bottom:1px solid #2a3040;color:#f0f2f5;font-size:15px;font-weight:bold;">
          ${logoHtml}${p.team}
        </td>
        <td style="padding:10px 16px;border-bottom:1px solid #2a3040;color:#c8a94a;font-size:15px;font-weight:bold;">${p.spread || ''}</td>
      </tr>`;
    }).join('');

    const name = data.teamName || data.name || data.email;

    let tiebreakerGameHtml = '';
    if (data.tiebreakerGame) {
      const parts = data.tiebreakerGame.split(' @ ');
      const awayAbbr = parts[0] ? parts[0].trim() : '';
      const homeAbbr = parts[1] ? parts[1].trim() : '';
      const awayLogo = logoUrl(awayAbbr);
      const homeLogo = logoUrl(homeAbbr);
      tiebreakerGameHtml =
        (awayLogo ? `<img src="${awayLogo}" width="24" height="24" style="vertical-align:middle;margin-right:4px;">` : '') +
        `<span style="color:#6b7588;font-size:14px;">${awayAbbr}</span>` +
        `<span style="color:#6b7588;font-size:13px;margin:0 8px;">@</span>` +
        (homeLogo ? `<img src="${homeLogo}" width="24" height="24" style="vertical-align:middle;margin-right:4px;">` : '') +
        `<span style="color:#6b7588;font-size:14px;">${homeAbbr}</span>`;
    }

    const tiebreakerSection = (data.tiebreaker !== undefined && data.tiebreaker !== '')
      ? `<div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;padding:16px;margin-bottom:20px;">
          <p style="color:#c8a94a;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin:0 0 10px;">Tiebreaker</p>
          ${tiebreakerGameHtml ? `<div style="margin-bottom:8px;">${tiebreakerGameHtml}</div>` : ''}
          <p style="color:#f0f2f5;font-size:22px;font-weight:bold;margin:0;">Total Score: <span style="color:#c8a94a;">${data.tiebreaker} pts</span></p>
        </div>`
      : '';

    const htmlBody = `
<div style="background:#0a0c10;padding:40px 20px;font-family:Arial,sans-serif;max-width:520px;margin:0 auto;">
  <div style="text-align:center;margin-bottom:32px;">
    <div style="font-size:36px;margin-bottom:8px;">🏈</div>
    <h1 style="color:#c8a94a;font-size:28px;letter-spacing:2px;margin:0;text-transform:uppercase;">Pick 5</h1>
    <p style="color:#6b7588;font-size:13px;letter-spacing:2px;text-transform:uppercase;margin:6px 0 0;">Picks Confirmation</p>
  </div>
  <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;padding:24px;margin-bottom:20px;">
    <p style="color:#6b7588;font-size:13px;letter-spacing:1px;text-transform:uppercase;margin:0 0 4px;">Hey ${name},</p>
    <p style="color:#f0f2f5;font-size:16px;margin:0;">Your picks for <strong style="color:#c8a94a;">${data.week}</strong> have been submitted!</p>
  </div>
  <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;overflow:hidden;margin-bottom:20px;">
    <div style="background:rgba(200,169,74,0.08);padding:12px 16px;border-bottom:1px solid #2a3040;">
      <p style="color:#c8a94a;font-size:13px;letter-spacing:2px;text-transform:uppercase;margin:0;">Your Picks</p>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr>
          <th style="padding:10px 16px;border-bottom:1px solid #2a3040;color:#6b7588;font-size:12px;letter-spacing:1px;text-align:left;text-transform:uppercase;">#</th>
          <th style="padding:10px 16px;border-bottom:1px solid #2a3040;color:#6b7588;font-size:12px;letter-spacing:1px;text-align:left;text-transform:uppercase;">Team</th>
          <th style="padding:10px 16px;border-bottom:1px solid #2a3040;color:#6b7588;font-size:12px;letter-spacing:1px;text-align:left;text-transform:uppercase;">Spread</th>
        </tr>
      </thead>
      <tbody>${picksList}</tbody>
    </table>
  </div>
  ${tiebreakerSection}
  <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;padding:16px;margin-bottom:24px;">
    <p style="color:#6b7588;font-size:12px;letter-spacing:1px;text-transform:uppercase;margin:0 0 4px;">Submitted</p>
    <p style="color:#f0f2f5;font-size:14px;margin:0;">${data.date || ''} at ${data.time || ''}</p>
  </div>
  <p style="color:#6b7588;font-size:12px;text-align:center;margin:0;">
    You can update your picks anytime before the games kick off.<br>
    Good luck this week! 🏆
  </p>
</div>`;

    MailApp.sendEmail({
      to: data.email,
      subject: '✅ Pick 5 — ' + data.week + ' Picks Confirmed',
      htmlBody: htmlBody
    });

  } catch(err) {
    Logger.log('Email failed: ' + err.toString());
  }
}

// ─── PICKS REMINDER ──────────────────────────────────────────────────────────

function sendPicksReminderEmail(data) {
  try {
    const week       = data.week       || 'this week';
    const recipients = data.recipients || [];

    recipients.forEach(recipient => {
      const toEmail = typeof recipient === 'string' ? recipient : (recipient.email || '');
      const name    = typeof recipient === 'string' ? recipient : (recipient.teamName || recipient.name || recipient.email || 'there');
      if (!toEmail) return;

      const htmlBody = `
<div style="background:#0a0c10;padding:36px 20px;font-family:Arial,sans-serif;max-width:520px;margin:0 auto;">
  <div style="text-align:center;margin-bottom:28px;">
    <div style="font-size:32px;margin-bottom:6px;">🏈</div>
    <h1 style="color:#c8a94a;font-size:26px;letter-spacing:2px;margin:0;text-transform:uppercase;">Pick 5</h1>
  </div>
  <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;padding:24px;margin-bottom:20px;">
    <p style="color:#6b7588;font-size:13px;letter-spacing:1px;text-transform:uppercase;margin:0 0 8px;">Hey ${name},</p>
    <p style="color:#f0f2f5;font-size:16px;margin:0 0 12px;">You haven't submitted your picks for <strong style="color:#c8a94a;">${week}</strong> yet!</p>
    <p style="color:#6b7588;font-size:14px;margin:0;">Log in and lock in your picks before the games kick off. Don't miss out! 🏆</p>
  </div>
  <div style="text-align:center;margin-bottom:20px;">
    <a href="${APP_URL}" style="background:#c8a94a;color:#0a0c10;padding:12px 28px;border-radius:6px;font-weight:bold;font-size:15px;text-decoration:none;display:inline-block;">Submit My Picks →</a>
  </div>
  <p style="color:#6b7588;font-size:12px;text-align:center;margin:0;">Good luck this week!</p>
</div>`;

      MailApp.sendEmail({
        to: toEmail,
        subject: '⏰ Pick 5 — Don\'t forget your ' + week + ' picks!',
        htmlBody: htmlBody
      });
    });

    Logger.log('Reminder sent to ' + recipients.length + ' players for ' + week);

  } catch(err) {
    Logger.log('Reminder failed: ' + err.toString());
    throw err;
  }
}

// ─── WEEKLY RECAP ────────────────────────────────────────────────────────────

function sendWeeklyRecapEmail(data) {
  try {
    const week          = data.week          || 'Week';
    const weekResults   = data.weekResults   || [];
    const leaderboard   = data.leaderboard   || [];
    const perfectScores = data.perfectScores || [];
    const shutouts      = data.shutouts      || [];
    const recipients    = data.recipients    || [];

    if (recipients.length === 0) return;

    let calloutsHtml = '';
    if (perfectScores.length > 0) {
      calloutsHtml += `<div style="background:linear-gradient(135deg,rgba(200,169,74,0.15),rgba(200,169,74,0.05));border:1px solid rgba(200,169,74,0.4);border-radius:8px;padding:14px 16px;margin-bottom:12px;text-align:center;">
        <span style="font-size:22px;">🔥</span>
        <span style="color:#c8a94a;font-size:14px;font-weight:bold;margin:0 10px;">${perfectScores.join(', ')}</span>
        <span style="color:#6b7588;font-size:13px;">went 5-0 this week!</span>
      </div>`;
    }
    if (shutouts.length > 0) {
      calloutsHtml += `<div style="background:rgba(224,85,85,0.06);border:1px solid rgba(224,85,85,0.3);border-radius:8px;padding:14px 16px;margin-bottom:12px;text-align:center;">
        <span style="font-size:22px;">💀</span>
        <span style="color:#e05555;font-size:14px;font-weight:bold;margin:0 10px;">${shutouts.join(', ')}</span>
        <span style="color:#6b7588;font-size:13px;">went 0-5 this week</span>
      </div>`;
    }

    const weekRows = weekResults.map((user, idx) => {
      const wins   = user.wins   !== undefined ? user.wins   : 0;
      const losses = user.losses !== undefined ? user.losses : 0;
      const recordColor = wins > losses ? '#4caf7d' : wins < losses ? '#e05555' : '#c8a94a';
      const isEven = idx % 2 === 1;
      let tbHtml = '<span style="color:#6b7588;">—</span>';
      if (user.tbScore !== null && user.tbScore !== undefined && user.tbGame) {
        const diffStr = user.tbDiff !== null && user.tbDiff !== undefined
          ? (user.tbDiff === 0
              ? ' <span style="color:#4caf7d;font-size:11px;">(exact!)</span>'
              : ` <span style="color:#6b7588;font-size:11px;">(off by ${user.tbDiff})</span>`)
          : '';
        tbHtml = `<strong style="color:#c8a94a;">${user.tbScore}</strong>${diffStr}`;
      }
      return `<tr style="background:${isEven ? 'rgba(255,255,255,0.02)' : 'transparent'};">
        <td style="padding:10px 14px;color:#6b7588;font-size:13px;width:28px;">${idx + 1}</td>
        <td style="padding:10px 14px;color:#f0f2f5;font-size:15px;font-weight:600;">${user.teamName || ''}</td>
        <td style="padding:10px 14px;text-align:center;">
          <span style="color:${recordColor};font-size:15px;font-weight:bold;">${wins}W – ${losses}L</span>
          ${user.pending > 0 ? `<br><span style="color:#6b7588;font-size:11px;">${user.pending} pending</span>` : ''}
        </td>
        <td style="padding:10px 14px;text-align:right;font-size:13px;">${tbHtml}</td>
      </tr>`;
    }).join('');

    const medalEmoji = ['🥇', '🥈', '🥉', '4️⃣', '5️⃣'];
    const leaderboardRows = leaderboard.map((entry) => {
      const rank        = entry.rank || 1;
      const medal       = rank <= 5 ? medalEmoji[rank - 1] : rank;
      const isTop       = rank === 1;
      const totalWins   = entry.totalWins   !== undefined ? entry.totalWins   : 0;
      const totalLosses = entry.totalLosses !== undefined ? entry.totalLosses : 0;
      return `<tr style="border-bottom:1px solid #1e2535;${isTop ? 'background:rgba(200,169,74,0.06);' : ''}">
        <td style="padding:10px 14px;font-size:16px;text-align:center;">${medal}</td>
        <td style="padding:10px 14px;color:#f0f2f5;font-size:14px;font-weight:${isTop ? 'bold' : 'normal'};">${entry.teamName || ''}</td>
        <td style="padding:10px 14px;text-align:right;color:${totalWins > totalLosses ? '#4caf7d' : '#f0f2f5'};font-size:14px;font-weight:bold;">${totalWins}W – ${totalLosses}L</td>
        <td style="padding:10px 14px;text-align:right;color:#6b7588;font-size:13px;">${entry.totalTbDiff !== null && entry.totalTbDiff !== undefined ? '+' + entry.totalTbDiff : ''}</td>
      </tr>`;
    }).join('');

    const htmlBody = `
<div style="background:#0a0c10;padding:36px 20px;font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">
  <div style="text-align:center;margin-bottom:28px;">
    <div style="font-size:32px;margin-bottom:6px;">🏈</div>
    <h1 style="color:#c8a94a;font-size:26px;letter-spacing:2px;margin:0;text-transform:uppercase;">Pick 5</h1>
    <p style="color:#f0f2f5;font-size:18px;font-weight:bold;margin:6px 0 0;">${week} Recap</p>
  </div>
  ${calloutsHtml}
  <div style="margin-bottom:28px;">
    <p style="color:#c8a94a;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin:0 0 12px;border-bottom:1px solid #2a3040;padding-bottom:8px;">This Week</p>
    <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;overflow:hidden;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:rgba(200,169,74,0.08);border-bottom:1px solid #2a3040;">
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">#</th>
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Team</th>
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:center;font-weight:normal;">Record</th>
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:right;font-weight:normal;">Tiebreaker</th>
          </tr>
        </thead>
        <tbody>${weekRows}</tbody>
      </table>
    </div>
  </div>
  <div>
    <p style="color:#c8a94a;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin:0 0 12px;border-bottom:1px solid #2a3040;padding-bottom:8px;">Season Standings</p>
    <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;overflow:hidden;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:rgba(200,169,74,0.08);border-bottom:1px solid #2a3040;">
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:center;font-weight:normal;width:40px;"></th>
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Team</th>
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:right;font-weight:normal;">Record</th>
            <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:right;font-weight:normal;">TB</th>
          </tr>
        </thead>
        <tbody>${leaderboardRows}</tbody>
      </table>
    </div>
  </div>
  <p style="color:#6b7588;font-size:12px;text-align:center;margin:24px 0 0;">Good luck next week! 🏆</p>
</div>`;

    // BCC-batch to stay within daily Gmail quota (1 send call per batch of 25)
    const toList = recipients.map(r => typeof r === 'string' ? r : (r.email || '')).filter(Boolean);
    const batchSize = 25;
    const failed = [];
    for (let i = 0; i < toList.length; i += batchSize) {
      const batch = toList.slice(i, i + batchSize);
      try {
        MailApp.sendEmail({
          to: ADMIN_EMAILS[0],
          bcc: batch.join(','),
          subject: '📊 Pick 5 — ' + week + ' Recap',
          htmlBody: htmlBody
        });
      } catch(e) {
        Logger.log('Recap batch failed (starting ' + i + '): ' + e);
        failed.push.apply(failed, batch);
      }
    }

    if (failed.length > 0) {
      MailApp.sendEmail({
        to: ADMIN_EMAILS.join(','),
        subject: '⚠️ Pick 5 Recap — ' + failed.length + ' emails not sent',
        body: 'The following recipients did not receive the ' + week + ' recap:\n\n' + failed.join('\n')
      });
    }

    Logger.log('Weekly recap: ' + (toList.length - failed.length) + ' sent, ' + failed.length + ' failed for ' + week);

  } catch(err) {
    Logger.log('Recap email failed: ' + err.toString());
    throw err;
  }
}

// ─── AUTOMATED REMINDER (time-based trigger) ─────────────────────────────────

function sendAutomatedPicksReminder() {
  var baseUrl = 'https://firestore.googleapis.com/v1/projects/' + FIREBASE_PROJECT_ID + '/databases/(default)/documents';

  var weekResp   = UrlFetchApp.fetch(baseUrl + '/settings/spreadsWeek?key=' + FIREBASE_API_KEY);
  var weekData   = JSON.parse(weekResp.getContentText());
  var activeWeek = weekData.fields && weekData.fields.value && weekData.fields.value.stringValue;
  if (!activeWeek) { Logger.log('No active week set — skipping.'); return; }
  var weekLabel = /^\d+$/.test(activeWeek) ? 'Week ' + activeWeek : activeWeek;

  var usersResp = UrlFetchApp.fetch(baseUrl + '/approvedUsers?key=' + FIREBASE_API_KEY + '&pageSize=300');
  var usersData = JSON.parse(usersResp.getContentText());
  var allUsers  = (usersData.documents || []).map(function(d) {
    var f = d.fields || {};
    return {
      email:    (f.email    && f.email.stringValue)    || '',
      teamName: (f.teamName && f.teamName.stringValue) || '',
      phone:    (f.phone    && f.phone.stringValue)    || '',
      isActive: f.isActive  ? f.isActive.booleanValue  : true
    };
  }).filter(function(u) { return u.email && u.isActive !== false; });

  var queryBody = {
    structuredQuery: {
      from: [{ collectionId: 'picks' }],
      where: { fieldFilter: { field: { fieldPath: 'week' }, op: 'EQUAL', value: { stringValue: weekLabel } } }
    }
  };
  var picksResp = UrlFetchApp.fetch(
    'https://firestore.googleapis.com/v1/projects/' + FIREBASE_PROJECT_ID + '/databases/(default)/documents:runQuery?key=' + FIREBASE_API_KEY,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(queryBody) }
  );
  var picksData = JSON.parse(picksResp.getContentText());

  var pickCountByEmail = {};
  (Array.isArray(picksData) ? picksData : []).forEach(function(result) {
    if (!result.document) return;
    var f     = result.document.fields || {};
    var email = ((f.email && f.email.stringValue) || '').toLowerCase();
    if (!email) return;
    var count = 0;
    ['pick1','pick2','pick3','pick4','pick5'].forEach(function(k) {
      if (f[k] && f[k].stringValue && f[k].stringValue.trim()) count++;
    });
    pickCountByEmail[email] = Math.max(pickCountByEmail[email] || 0, count);
  });

  var noPicksUsers = [], partialUsers = [];
  allUsers.forEach(function(u) {
    var count = pickCountByEmail[u.email.toLowerCase()] || 0;
    if (count === 0)    noPicksUsers.push({ email: u.email, teamName: u.teamName, phone: u.phone, count: 0 });
    else if (count < 5) partialUsers.push({ email: u.email, teamName: u.teamName, phone: u.phone, count: count });
  });

  Logger.log('Week: ' + weekLabel + ' | No picks: ' + noPicksUsers.length + ' | Partial: ' + partialUsers.length);
  noPicksUsers.forEach(function(u) { sendReminderEmail(u.email, u.teamName, weekLabel, 0); });
  partialUsers.forEach(function(u) { sendReminderEmail(u.email, u.teamName, weekLabel, u.count); });
  sendAdminSummaryEmail(weekLabel, noPicksUsers, partialUsers, allUsers);
}

function sendReminderEmail(email, teamName, weekLabel, pickCount) {
  var name = teamName || 'there';
  var subject, bodyHtml;

  if (pickCount === 0) {
    subject  = '⏰ Reminder: Submit your picks for ' + weekLabel + '!';
    bodyHtml = '<p style="color:#f0f2f5;font-size:16px;margin:0 0 12px;">You haven\'t submitted any picks for <strong style="color:#c8a94a;">' + weekLabel + '</strong> yet!</p>'
             + '<p style="color:#6b7588;font-size:14px;margin:0;">Get in and lock them in before the games kick off.</p>';
  } else {
    var remaining = 5 - pickCount;
    subject  = '⏰ Reminder: You still need ' + remaining + ' more pick' + (remaining !== 1 ? 's' : '') + ' for ' + weekLabel + '!';
    bodyHtml = '<p style="color:#f0f2f5;font-size:16px;margin:0 0 12px;">You\'ve submitted <strong style="color:#c8a94a;">' + pickCount + ' out of 5</strong> picks for ' + weekLabel + '.</p>'
             + '<p style="color:#6b7588;font-size:14px;margin:0;">You\'re not done yet! Get back in and lock in your remaining ' + remaining + ' pick' + (remaining !== 1 ? 's' : '') + '.</p>';
  }

  var htmlBody = `
<div style="background:#0a0c10;padding:36px 20px;font-family:Arial,sans-serif;max-width:520px;margin:0 auto;">
  <div style="text-align:center;margin-bottom:28px;">
    <div style="font-size:32px;margin-bottom:6px;">🏈</div>
    <h1 style="color:#c8a94a;font-size:26px;letter-spacing:2px;margin:0;text-transform:uppercase;">Pick 5</h1>
  </div>
  <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;padding:24px;margin-bottom:20px;">
    <p style="color:#6b7588;font-size:13px;letter-spacing:1px;text-transform:uppercase;margin:0 0 8px;">Hey ${name},</p>
    ${bodyHtml}
  </div>
  <div style="text-align:center;margin-bottom:20px;">
    <a href="${APP_URL}" style="background:#c8a94a;color:#0a0c10;padding:12px 28px;border-radius:6px;font-weight:bold;font-size:15px;text-decoration:none;display:inline-block;">Submit My Picks →</a>
  </div>
  <p style="color:#6b7588;font-size:12px;text-align:center;margin:0;">Good luck this week! 🏆</p>
</div>`;

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
  Logger.log('Sent to ' + email + ' (' + pickCount + '/5 picks)');
}

function sendAdminSummaryEmail(weekLabel, noPicksUsers, partialUsers, allUsers) {
  var subject, htmlBody;

  if (noPicksUsers.length === 0 && partialUsers.length === 0) {
    subject = '✅ Pick5 ' + weekLabel + ' — Everyone has submitted!';
    htmlBody = `
<div style="background:#0a0c10;padding:36px 20px;font-family:Arial,sans-serif;max-width:540px;margin:0 auto;">
  <div style="text-align:center;margin-bottom:28px;">
    <div style="font-size:32px;margin-bottom:6px;">🏈</div>
    <h1 style="color:#c8a94a;font-size:22px;letter-spacing:2px;margin:0;text-transform:uppercase;">Pick 5 Admin</h1>
    <p style="color:#6b7588;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin:6px 0 0;">${weekLabel} Reminder Summary</p>
  </div>
  <div style="background:linear-gradient(135deg,rgba(76,175,125,0.12),rgba(76,175,125,0.04));border:1px solid rgba(76,175,125,0.4);border-radius:8px;padding:20px 24px;text-align:center;">
    <div style="font-size:28px;margin-bottom:8px;">🎉</div>
    <p style="color:#4caf7d;font-size:18px;font-weight:bold;margin:0 0 6px;">All ${allUsers.length} players have submitted!</p>
    <p style="color:#6b7588;font-size:14px;margin:0;">Nothing to do — everyone is locked in for ${weekLabel}.</p>
  </div>
</div>`;
    MailApp.sendEmail({ to: ADMIN_EMAILS.join(','), subject: subject, htmlBody: htmlBody });
    return;
  }

  function buildRows(users, showCount) {
    return users.map(function(u) {
      var countCell = showCount
        ? `<td style="padding:10px 14px;text-align:right;"><span style="color:#c8a94a;font-weight:bold;">${u.count}/5</span></td>`
        : `<td style="padding:10px 14px;text-align:right;"><span style="color:#e05555;font-weight:bold;">0/5</span></td>`;
      return `<tr style="border-bottom:1px solid #1e2535;">
        <td style="padding:10px 14px;color:#f0f2f5;font-size:14px;font-weight:600;">${u.teamName || '—'}</td>
        <td style="padding:10px 14px;color:#6b7588;font-size:13px;">${u.email}</td>
        <td style="padding:10px 14px;color:#6b7588;font-size:13px;">${u.phone || '—'}</td>
        ${countCell}
      </tr>`;
    }).join('');
  }

  var noPicksSection = '';
  if (noPicksUsers.length > 0) {
    noPicksSection = `
    <div style="margin-bottom:24px;">
      <p style="color:#e05555;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin:0 0 10px;border-bottom:1px solid #2a3040;padding-bottom:8px;">
        💀 No Picks Submitted (${noPicksUsers.length})
      </p>
      <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;overflow:hidden;">
        <table style="width:100%;border-collapse:collapse;">
          <thead>
            <tr style="background:rgba(224,85,85,0.06);border-bottom:1px solid #2a3040;">
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Team</th>
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Email</th>
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Phone</th>
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:right;font-weight:normal;">Picks</th>
            </tr>
          </thead>
          <tbody>${buildRows(noPicksUsers, false)}</tbody>
        </table>
      </div>
    </div>`;
  }

  var partialSection = '';
  if (partialUsers.length > 0) {
    partialSection = `
    <div style="margin-bottom:24px;">
      <p style="color:#c8a94a;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin:0 0 10px;border-bottom:1px solid #2a3040;padding-bottom:8px;">
        ⚠️ Partial Picks (${partialUsers.length})
      </p>
      <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;overflow:hidden;">
        <table style="width:100%;border-collapse:collapse;">
          <thead>
            <tr style="background:rgba(200,169,74,0.06);border-bottom:1px solid #2a3040;">
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Team</th>
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Email</th>
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:left;font-weight:normal;">Phone</th>
              <th style="padding:8px 14px;color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;text-align:right;font-weight:normal;">Picks</th>
            </tr>
          </thead>
          <tbody>${buildRows(partialUsers, true)}</tbody>
        </table>
      </div>
    </div>`;
  }

  var total = noPicksUsers.length + partialUsers.length;
  subject = '⏰ Pick5 ' + weekLabel + ' — ' + total + ' player' + (total !== 1 ? 's' : '') + ' need a nudge';
  htmlBody = `
<div style="background:#0a0c10;padding:36px 20px;font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">
  <div style="text-align:center;margin-bottom:28px;">
    <div style="font-size:32px;margin-bottom:6px;">🏈</div>
    <h1 style="color:#c8a94a;font-size:22px;letter-spacing:2px;margin:0;text-transform:uppercase;">Pick 5 Admin</h1>
    <p style="color:#6b7588;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin:6px 0 0;">${weekLabel} Reminder Summary</p>
  </div>
  <div style="background:#13171f;border:1px solid #2a3040;border-radius:8px;padding:14px 20px;margin-bottom:24px;display:flex;justify-content:space-between;gap:12px;text-align:center;">
    <div style="flex:1;">
      <p style="color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;margin:0 0 4px;">Total Players</p>
      <p style="color:#f0f2f5;font-size:22px;font-weight:bold;margin:0;">${allUsers.length}</p>
    </div>
    <div style="flex:1;border-left:1px solid #2a3040;">
      <p style="color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;margin:0 0 4px;">No Picks</p>
      <p style="color:#e05555;font-size:22px;font-weight:bold;margin:0;">${noPicksUsers.length}</p>
    </div>
    <div style="flex:1;border-left:1px solid #2a3040;">
      <p style="color:#6b7588;font-size:11px;letter-spacing:1px;text-transform:uppercase;margin:0 0 4px;">Partial</p>
      <p style="color:#c8a94a;font-size:22px;font-weight:bold;margin:0;">${partialUsers.length}</p>
    </div>
  </div>
  ${noPicksSection}
  ${partialSection}
  <p style="color:#6b7588;font-size:12px;text-align:center;margin:0;">Reminder emails have been sent to all ${total} player${total !== 1 ? 's' : ''} above.</p>
</div>`;

  MailApp.sendEmail({ to: ADMIN_EMAILS.join(','), subject: subject, htmlBody: htmlBody });
}

// ─── ONE-TIME EXPORT ─────────────────────────────────────────────────────────

function exportAllPicksToSheet() {
  var baseUrl = 'https://firestore.googleapis.com/v1/projects/' + FIREBASE_PROJECT_ID + '/databases/(default)/documents';
  var allDocs = [];
  var pageToken = '';

  // Paginate through all picks documents
  do {
    var url = baseUrl + '/picks?key=' + FIREBASE_API_KEY + '&pageSize=300';
    if (pageToken) url += '&pageToken=' + encodeURIComponent(pageToken);
    var resp = UrlFetchApp.fetch(url);
    var data = JSON.parse(resp.getContentText());
    if (data.documents) allDocs = allDocs.concat(data.documents);
    pageToken = data.nextPageToken || '';
  } while (pageToken);

  Logger.log('Total picks documents fetched: ' + allDocs.length);

  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Picks') || ss.insertSheet('Picks');

  // Clear and write header
  sheet.clearContents();
  sheet.appendRow(['Timestamp', 'Team Name', 'Email', 'Week',
                   'Pick 1', 'Pick 2', 'Pick 3', 'Pick 4', 'Pick 5',
                   'Spread 1', 'Spread 2', 'Spread 3', 'Spread 4', 'Spread 5',
                   'Tiebreaker', 'Tiebreaker Game']);

  // Sort: by week then team name
  var rows = allDocs.map(function(doc) {
    var f = doc.fields || {};
    function str(k) { return (f[k] && f[k].stringValue) || ''; }
    function num(k) { return (f[k] && (f[k].integerValue || f[k].doubleValue)) || ''; }

    var ts = str('submittedAt') || str('timestamp') || '';
    var tbVal = str('tiebreaker') || num('tiebreaker') || '';

    return [
      ts, str('teamName') || str('name'), str('email'), str('week'),
      str('pick1'), str('pick2'), str('pick3'), str('pick4'), str('pick5'),
      str('spread1'), str('spread2'), str('spread3'), str('spread4'), str('spread5'),
      tbVal, str('tiebreakerGame')
    ];
  });

  rows.sort(function(a, b) {
    var wA = a[3] || '', wB = b[3] || '';
    if (wA < wB) return -1;
    if (wA > wB) return 1;
    return (a[1] || '').localeCompare(b[1] || '');
  });

  rows.forEach(function(row) { sheet.appendRow(row); });

  Logger.log('Export complete. ' + rows.length + ' rows written to Picks sheet.');
  SpreadsheetApp.flush();
}

// ─── TEST FUNCTIONS ───────────────────────────────────────────────────────────

function testConfirmation() {
  sendPicksConfirmation({
    email: 'jenowakoski@gmail.com', teamName: 'Iksokawon',
    date: new Date().toLocaleDateString(), time: new Date().toLocaleTimeString(),
    week: 'Week 2',
    pick1: 'KC',  spread1: '-3',   pick2: 'PHI', spread2: '-5.5',
    pick3: 'BUF', spread3: '-1.5', pick4: 'DET', spread4: '-7',
    pick5: 'PIT', spread5: '-3.5',
    tiebreaker: 44, tiebreakerGame: 'DEN @ KC'
  });
}

function testReminder() {
  sendPicksReminderEmail({
    week: 'Week 2',
    recipients: [{ teamName: 'Iksokawon', email: 'jenowakoski@gmail.com' }]
  });
}

function testRecap() {
  sendWeeklyRecapEmail({
    week: 'Week 1',
    recipients: ['jenowakoski@gmail.com'],
    weekResults: [
      { teamName: 'Breezy',  wins: 4, losses: 1, tbScore: 42, tbGame: 'DEN @ KC', tbDiff: 1, pending: 0 },
      { teamName: 'Commish', wins: 3, losses: 2, tbScore: 45, tbGame: 'DEN @ KC', tbDiff: 4, pending: 0 },
      { teamName: 'Doesnot', wins: 0, losses: 5, tbScore: 47, tbGame: 'DEN @ KC', tbDiff: 6, pending: 0 },
    ],
    leaderboard: [
      { rank: 1, teamName: 'Breezy',  totalWins: 4, totalLosses: 1, totalTbDiff: 1 },
      { rank: 2, teamName: 'Commish', totalWins: 3, totalLosses: 2, totalTbDiff: 4 },
      { rank: 3, teamName: 'Doesnot', totalWins: 0, totalLosses: 5, totalTbDiff: 6 },
    ],
    perfectScores: [],
    shutouts: ['Doesnot']
  });
}

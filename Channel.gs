/* ══════════════════════════════════════════════════════════════════
   КАНАЛ TELEGRAM — серверная часть (новый файл Channel.gs)
   Берёт статистику канала @di_ero из Supabase: таблицы posts и channel_daily.
   Саму статистику собирает Supabase раз в 3 часа, здесь только чтение.

   Ключ ниже — «публичный» (anon): так задумано Supabase, его можно видеть всем.
   Таблицы защищены правилами: читать можно, записывать нельзя.
══════════════════════════════════════════════════════════════════ */

var CHANNEL_SB_URL = 'https://hedypadnuufbpeoqvbyu.supabase.co';
var CHANNEL_SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhlZHlwYWRudXVmYnBlb3F2Ynl1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTM4NDE2NTIsImV4cCI6MjEwOTQxNzY1Mn0.TurwiYmOADlLjGZU2Ocq_shY_0Tjyu9TIyG1qcUUgU4';

// Вызывается из страницы: Api.call('pultChannelStats')
function pultChannelStats() {
  function get(path) { // один запрос к Supabase
    var r = UrlFetchApp.fetch(CHANNEL_SB_URL + '/rest/v1/' + path, {
      headers: { apikey: CHANNEL_SB_KEY, Authorization: 'Bearer ' + CHANNEL_SB_KEY },
      muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) {
      throw new Error('Supabase ответил ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
    }
    return JSON.parse(r.getContentText());
  }

  var posts = get('posts?select=post_id,body,posted_at,media_kind,views,reactions,reactions_total,url,updated_at&order=posted_at.desc&limit=1000');
  posts.forEach(function (p) { if (p.body) p.body = p.body.slice(0, 400); }); // длинные тексты не тащим целиком

  return {
    ok: true,
    posts: posts,
    daily: get('channel_daily?select=day,subscribers&order=day.desc&limit=60')
  };
}

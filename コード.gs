// =================================================================
// ⚙️ 設定エリア：キー・トークン・IDはスクリプト プロパティから読み込みます
// =================================================================
//
// 「プロジェクトの設定」→「スクリプト プロパティ」に以下のキーを登録してください
//
// DROPBOX_APP_KEY         Dropbox App Consoleの「App key」
// DROPBOX_APP_SECRET      Dropbox App Consoleの「App secret」
// DROPBOX_REFRESH_TOKEN   初回セットアップ(step2)で自動保存されます（手動登録は不要）
// ACCOUNT_1_TOKEN         1つ目のBufferアカウントのトークン
// ACCOUNT_1_PROFILE_ID    1つ目のBufferアカウントのチャンネルID（カンマ区切り）
// ACCOUNT_2_TOKEN         2つ目のBufferアカウントのトークン（使わない場合は未登録でOK）
// ACCOUNT_2_PROFILE_ID    2つ目のBufferアカウントのチャンネルID（カンマ区切り）
// BLOGGER_BLOG_ID         BloggerのブログID（未登録の場合はBlogger投稿をスキップ）
//
// ※ BLOGGER_POST_COUNTER（連番）は、これまで通りスクリプトが自動で管理します

const PROPS = PropertiesService.getScriptProperties();

function getProp_(key) {
  return PROPS.getProperty(key) || '';
}

const CONFIG = {
  // 【1】Dropbox App Consoleの「App key」「App secret」
  DROPBOX_APP_KEY: getProp_('DROPBOX_APP_KEY'),
  DROPBOX_APP_SECRET: getProp_('DROPBOX_APP_SECRET'),

  // 【2】リフレッシュトークン（初回のみ、下の「初回セットアップ用の関数」で取得）
  // アクセストークンは4時間で失効するため、毎回このリフレッシュトークンから
  // 新しいアクセストークンを自動生成する仕組みにしています。
  DROPBOX_REFRESH_TOKEN: getProp_('DROPBOX_REFRESH_TOKEN'),

  // 【3】Dropbox内の「Buffer投稿用」フォルダのパス
  // App Folder方式（App Folder: MIYOKO_NEWS）の場合、そのフォルダ自体がAPI上のルート"/"として
  // 扱われるので、直下に置くなら空文字 '' でOKです。サブフォルダを作った場合のみパスを指定してください（例: '/2026年'）。
  DROPBOX_FOLDER_PATH: '',

  // 【3】1つ目のBufferアカウントの情報
  ACCOUNT_1: {
    TOKEN: getProp_('ACCOUNT_1_TOKEN'),
    PROFILE_ID: getProp_('ACCOUNT_1_PROFILE_ID')
  },

  // 【4】2つ目のBufferアカウントの情報（使わない場合は未登録でOK）
  ACCOUNT_2: {
    TOKEN: getProp_('ACCOUNT_2_TOKEN'),
    PROFILE_ID: getProp_('ACCOUNT_2_PROFILE_ID')
  },

  // 📝 Blogger（API投稿）の設定
  // ※ メール投稿(Mail-to-Blogger)は、そもそも「ラベルを付けられない」「画像サイズを指定できない」
  //    という仕様上の制限があるため、Blogger APIを直接呼び出す方式に変更しました。
  BLOGGER_BLOG_ID: getProp_('BLOGGER_BLOG_ID'), // 下の debug_getBlogIdFromUrl() で調べて登録してください
  BLOGGER_LABELS: ['leblanc', 'ルブラン美夜子'], // 付けたいラベル（タグ）を配列で指定
  BLOGGER_TITLE: 'ルブラン美夜子のレッツウェザー#', // 投稿タイトル（固定文言。末尾に自動で連番が付きます）
  BLOGGER_POST_NUMBER_START: 742, // 連番の開始値（初回投稿がこの数字になります）
  BLOGGER_IMAGE_MAX_WIDTH_PX: 1600 // ブログ内での画像の最大表示幅（px）。大きくするほど「特大」表示に近づきます

};

// スクリプト プロパティの登録漏れを確認する（手動実行用）
function checkScriptProperties() {
  const required = [
    'DROPBOX_APP_KEY',
    'DROPBOX_APP_SECRET',
    'DROPBOX_REFRESH_TOKEN',
    'ACCOUNT_1_TOKEN',
    'ACCOUNT_1_PROFILE_ID'
  ];
  const optional = [
    'ACCOUNT_2_TOKEN',
    'ACCOUNT_2_PROFILE_ID',
    'BLOGGER_BLOG_ID'
  ];

  const missingRequired = required.filter(key => !getProp_(key));
  const missingOptional = optional.filter(key => !getProp_(key));

  if (missingRequired.length === 0) {
    Logger.log('✅ 必須のスクリプト プロパティはすべて登録されています。');
  } else {
    Logger.log('❌ 未登録の必須プロパティ: ' + missingRequired.join(', '));
  }

  if (missingOptional.length > 0) {
    Logger.log('ℹ️ 未登録の任意プロパティ（未登録の場合は該当の投稿をスキップします）: ' + missingOptional.join(', '));
  }
}

// =================================================================
// 🚀 メイン処理：毎朝自動で実行されるプログラム
// =================================================================

function checkAndPostToBufferAndBlogger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    const rowDateStr = Utilities.formatDate(new Date(data[i][0]), "Asia/Tokyo", "yyyy/MM/dd");
    const text = data[i][1];
    const fileName = data[i][2];
    const status = data[i][3];

    if (rowDateStr === todayStr && !status) {
      Logger.log(`▶️ 行 ${i + 1} の処理を開始します（ファイル名: ${fileName}）`);

      try {
        let finalMediaUrl = null;

        const isVideo = checkIsVideo(fileName);

        if (fileName && fileName.trim() !== "") {
          // 💡 【ここをDropbox経由に変更】
          // リフレッシュトークンから、有効なアクセストークンをその都度取得
          Logger.log('🔑 Dropboxのアクセストークンを取得中...');
          const dropboxAccessToken = getDropboxAccessToken();
          Logger.log('🔑 Dropboxのアクセストークン取得完了');

          const dropboxPath = joinDropboxPath(CONFIG.DROPBOX_FOLDER_PATH, fileName);
          Logger.log(`📂 Dropboxパス: ${dropboxPath}`);

          // ファイルが実際に存在するか軽量チェック（ダウンロードはしない）
          if (!dropboxFileExists(dropboxAccessToken, dropboxPath)) {
            Logger.log('❌ Dropboxにファイルが見つかりません');
            sheet.getRange(i + 1, 4).setValue("エラー: Dropboxにファイルが見つかりません");
            continue;
          }
          Logger.log('✅ Dropboxファイルの存在確認OK');

          // Buffer・Blogger用の直リンクURLを取得（共有リンクをraw=1形式に変換）
          finalMediaUrl = getDropboxDirectUrl(dropboxAccessToken, dropboxPath);

          if (!finalMediaUrl) {
            Logger.log('❌ Dropbox共有リンクの取得に失敗しました');
            sheet.getRange(i + 1, 4).setValue("エラー: Dropboxの共有リンク取得に失敗しました");
            continue;
          }
          Logger.log(`✅ Dropbox直リンク取得成功: ${finalMediaUrl}`);

        } else {
          sheet.getRange(i + 1, 4).setValue("エラー: 全てのポストに画像（動画）が必要です");
          continue;
        }

        // ① Bufferへの予約投稿処理
        Logger.log('📤 Bufferアカウント1への投稿を開始します...');
        const res1 = postToEachChannels(CONFIG.ACCOUNT_1.TOKEN, CONFIG.ACCOUNT_1.PROFILE_ID, text, finalMediaUrl, isVideo);
        Logger.log(`📤 Bufferアカウント1への投稿結果: ${res1}`);
        Utilities.sleep(2000);

        let res2 = true;
        if (CONFIG.ACCOUNT_2.TOKEN) {
          Logger.log('📤 Bufferアカウント2への投稿を開始します...');
          res2 = postToEachChannels(CONFIG.ACCOUNT_2.TOKEN, CONFIG.ACCOUNT_2.PROFILE_ID, text, finalMediaUrl, isVideo);
          Logger.log(`📤 Bufferアカウント2への投稿結果: ${res2}`);
        } else {
          Logger.log('⏭️ Bufferアカウント2はスキップ（TOKEN未設定）');
        }

        // ② BloggerへのAPI投稿処理（ラベル・画像サイズを指定できる）
        let bloggerRes = false;
        if (CONFIG.BLOGGER_BLOG_ID) {
          Logger.log(`📝 BloggerへのAPI投稿を開始します（Blog ID: ${CONFIG.BLOGGER_BLOG_ID}）...`);
          bloggerRes = postToBlogger(text, finalMediaUrl, isVideo);
          Logger.log(`📝 BloggerへのAPI投稿結果: ${bloggerRes}`);
        } else {
          Logger.log('⏭️ Blogger投稿はスキップ（BLOGGER_BLOG_ID未設定）');
        }

        if (res1 && res2 && bloggerRes) {
          sheet.getRange(i + 1, 4).setValue("完了（Buffer・Blogger）");
          Logger.log(`行 ${i + 1}: すべての処理が正常に完了しました！`);
        } else {
          Logger.log(`⚠️ 一部処理が未完了です（res1=${res1}, res2=${res2}, bloggerRes=${bloggerRes}）。ステータス欄は更新されません。`);
        }

      } catch (error) {
        sheet.getRange(i + 1, 4).setValue(`エラー: ${error.message}`);
        Logger.log(`❌ 行 ${i + 1} でエラー発生: ${error.message}`);
      }
      break;
    }
  }
}

// =================================================================
// 📝 Blogger連携用の関数群（Blogger APIを直接呼び出す方式）
// =================================================================
//
// 【事前準備】
// 1. appsscript.json の oauthScopes に "https://www.googleapis.com/auth/blogger" が
//    含まれていればOK（Advanced Serviceの追加は不要です）
// 2. スクリプト プロパティ BLOGGER_BLOG_ID に、あなたのブログIDを登録
//    （下の debug_getBlogIdFromUrl() 関数で調べられます）
// 3. 初回実行時、Googleアカウントでの権限承認画面が出るので許可してください

// 現在のユーザーの権限でBlogger APIを呼ぶためのOAuthトークンを取得する
function getBloggerAuthToken() {
  return ScriptApp.getOAuthToken();
}

// タイトル末尾に付ける連番を取得し、1つ進めて保存する
// （PropertiesServiceに記憶させているので、実行のたびに増え続けます）
function getNextPostNumber() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('BLOGGER_POST_COUNTER');
  const nextNumber = stored ? parseInt(stored, 10) + 1 : CONFIG.BLOGGER_POST_NUMBER_START;
  props.setProperty('BLOGGER_POST_COUNTER', String(nextNumber));
  return nextNumber;
}

// デバッグ用: 現在保存されている連番の値と、次回投稿される番号を確認する
function debug_checkPostNumber() {
  const stored = PropertiesService.getScriptProperties().getProperty('BLOGGER_POST_COUNTER');
  Logger.log(`現在保存されている値: ${stored}`);
  Logger.log(`次回投稿される番号: ${stored ? parseInt(stored, 10) + 1 : CONFIG.BLOGGER_POST_NUMBER_START}`);
}

// デバッグ用: 連番を任意の数字にリセットしたい場合に使う
function debug_resetPostNumber() {
  const resetTo = CONFIG.BLOGGER_POST_NUMBER_START - 1; // 次回実行時にSTART値から始まる
  PropertiesService.getScriptProperties().setProperty('BLOGGER_POST_COUNTER', String(resetTo));
  Logger.log(`✅ 連番をリセットしました。次回投稿は #${CONFIG.BLOGGER_POST_NUMBER_START} になります`);
}

function postToBlogger(text, mediaUrl, isVideo) {
  const htmlBody = buildBloggerHtmlBody(text, mediaUrl, isVideo);
  const title = CONFIG.BLOGGER_TITLE + getNextPostNumber();

  const post = {
    kind: 'blogger#post',
    title: title,
    content: htmlBody,
    labels: CONFIG.BLOGGER_LABELS
  };

  const url = `https://www.googleapis.com/blogger/v3/blogs/${CONFIG.BLOGGER_BLOG_ID}/posts/`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + getBloggerAuthToken() },
    payload: JSON.stringify(post),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 200 || code === 201) {
    const created = JSON.parse(response.getContentText());
    Logger.log(`✅ Blogger投稿が作成されました: ${created.url}`);
    return true;
  } else {
    throw new Error(`Blogger投稿に失敗しました: HTTP ${code} - ${response.getContentText()}`);
  }
}

// Blogger投稿用のHTML本文を組み立てる（画像・動画を「特大」表示できるサイズで埋め込む）
function buildBloggerHtmlBody(text, mediaUrl, isVideo) {
  const escapedText = text.replace(/\n/g, '<br>');
  const maxWidth = CONFIG.BLOGGER_IMAGE_MAX_WIDTH_PX;

  let mediaHtml;
  if (isVideo) {
    mediaHtml = `<video controls style="width:100%; max-width:${maxWidth}px; height:auto; display:block; margin:0 auto;">
  <source src="${mediaUrl}">
</video>`;
  } else {
    mediaHtml = `<img src="${mediaUrl}" style="width:100%; max-width:${maxWidth}px; height:auto; display:block; margin:0 auto;" />`;
  }

  return `${escapedText}<br>${mediaHtml}`;
}

// デバッグ用: ブログURLからBlog IDを調べる
function debug_getBlogIdFromUrl() {
  const blogUrl = 'ここに実際のブログURLを入れる（例: https://xxxx.blogspot.com/）';
  const url = 'https://www.googleapis.com/blogger/v3/blogs/byurl?url=' + encodeURIComponent(blogUrl);
  const options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + getBloggerAuthToken() },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 200) {
    const blog = JSON.parse(response.getContentText());
    Logger.log(`Blog ID: ${blog.id}`);
    Logger.log(`Blog名: ${blog.name}`);
  } else {
    Logger.log(`❌ 取得失敗: HTTP ${code} - ${response.getContentText()}`);
  }
}

function checkIsVideo(fileName) {
  if (!fileName) return false;
  const lowerName = fileName.toLowerCase();
  return lowerName.endsWith('.mp4') || lowerName.endsWith('.mov') || lowerName.endsWith('.avi') || lowerName.endsWith('.webm') || lowerName.endsWith('.mkv');
}

// ファイル名の拡張子から正しいMIMEタイプを判定する
// （Dropboxのdownload APIはContent-Typeをapplication/octet-streamで返すことが多く、
//   そのままメールに添付するとBloggerが画像として認識しないことがあるため）
function getMimeTypeFromFileName(fileName) {
  const lowerName = fileName.toLowerCase();
  const mimeMap = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
    '.mp4': 'video/mp4',
    '.mov': 'video/quicktime',
    '.avi': 'video/x-msvideo',
    '.webm': 'video/webm',
    '.mkv': 'video/x-matroska'
  };
  for (const ext in mimeMap) {
    if (lowerName.endsWith(ext)) return mimeMap[ext];
  }
  return 'application/octet-stream';
}

// =================================================================
// 📦 Dropbox連携用の関数群
// =================================================================

// リフレッシュトークンを使って、有効期限内の新しいアクセストークンを取得する
// （アクセストークンは4時間で失効するため、実行のたびに呼び出す）
// 💡 502/503/504/500などの一時的なサーバーエラーは、少し待って自動的に再試行する
function getDropboxAccessToken() {
  const url = 'https://api.dropboxapi.com/oauth2/token';
  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: CONFIG.DROPBOX_REFRESH_TOKEN,
      client_id: CONFIG.DROPBOX_APP_KEY,
      client_secret: CONFIG.DROPBOX_APP_SECRET
    },
    muteHttpExceptions: true
  };

  const maxRetries = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    let response;
    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (networkError) {
      lastError = networkError;
      if (attempt < maxRetries) {
        Logger.log(`⏳ Dropboxへの接続に失敗しました(${networkError.message})。${attempt}回目の再試行前に待機します...`);
        Utilities.sleep(3000 * attempt);
        continue;
      } else {
        throw new Error(`Dropboxへの接続に${maxRetries}回失敗しました: ${networkError.message}`);
      }
    }

    const code = response.getResponseCode();

    if (code === 200) {
      const json = JSON.parse(response.getContentText());
      return json.access_token;
    }

    if ([500, 502, 503, 504].includes(code) && attempt < maxRetries) {
      Logger.log(`⏳ Dropboxが一時的なエラー(HTTP ${code})を返しました。${attempt}回目の再試行前に待機します...`);
      Utilities.sleep(3000 * attempt);
      lastError = new Error(`HTTP ${code} - ${response.getContentText()}`);
      continue;
    }

    throw new Error(`Dropboxアクセストークンの更新に失敗しました: HTTP ${code} - ${response.getContentText()}`);
  }

  throw lastError;
}

// フォルダパスとファイル名を安全に連結する（先頭の"/"漏れ・末尾の"/"重複を自動で補正）
function joinDropboxPath(folderPath, fileName) {
  let base = (folderPath || '').trim();
  base = base.replace(/\/+$/, '');       // 末尾の "/" を除去
  if (base && !base.startsWith('/')) {
    base = '/' + base;                    // 先頭に "/" が無ければ補う
  }
  return base + '/' + fileName.trim();
}

// 指定パスのファイルが存在するか、軽量なメタデータ取得で確認する（ダウンロードはしない）
function dropboxFileExists(accessToken, filePath) {
  const url = 'https://api.dropboxapi.com/2/files/get_metadata';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + accessToken },
    payload: JSON.stringify({ path: filePath }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  return response.getResponseCode() === 200;
}

// 指定パスのファイルをDropboxからダウンロードしてBlobとして返す（現在は未使用・デバッグ用に残置）
function dropboxDownloadFile(accessToken, filePath) {
  const url = 'https://content.dropboxapi.com/2/files/download';
  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'Dropbox-API-Arg': JSON.stringify({ path: filePath })
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 200) {
    return response.getBlob();
  } else {
    Logger.log(`Dropboxダウンロード失敗 (${filePath}): HTTP ${code} - ${response.getContentText()}`);
    return null;
  }
}

// 指定パスの「外部から直接アクセスできるURL」を取得する
// 既存の共有リンクがあればそれを再利用し、無ければ新規作成する
function getDropboxDirectUrl(accessToken, filePath) {
  let sharedUrl = dropboxListExistingSharedLink(accessToken, filePath);

  if (!sharedUrl) {
    sharedUrl = dropboxCreateSharedLink(accessToken, filePath);
  }

  if (!sharedUrl) return null;

  // dl.dropbox.com のドメインに変換し、raw=1 を付与することで、
  // 一部のサービス（X/Twitterなど）がリダイレクトを追従できず失敗する問題を回避する
  let directUrl = sharedUrl
    .replace('www.dropbox.com', 'dl.dropboxusercontent.com')
    .replace('?dl=0', '?raw=1')
    .replace('&dl=0', '&raw=1');

  if (!directUrl.includes('raw=1')) {
    directUrl += (directUrl.includes('?') ? '&' : '?') + 'raw=1';
  }

  return directUrl;
}

// 既存の共有リンクを検索する
function dropboxListExistingSharedLink(accessToken, filePath) {
  const url = 'https://api.dropboxapi.com/2/sharing/list_shared_links';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + accessToken },
    payload: JSON.stringify({ path: filePath, direct_only: true }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    const json = JSON.parse(response.getContentText());
    if (json.links && json.links.length > 0) {
      return json.links[0].url;
    }
  }
  return null;
}

// 新規に共有リンクを作成する
function dropboxCreateSharedLink(accessToken, filePath) {
  const url = 'https://api.dropboxapi.com/2/sharing/create_shared_link_with_settings';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + accessToken },
    payload: JSON.stringify({ path: filePath }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 200) {
    const json = JSON.parse(response.getContentText());
    return json.url;
  } else {
    // 既にリンクが存在する場合のエラーはlist_shared_linksで拾えるはずだが、念のためログを残す
    Logger.log(`Dropbox共有リンク作成失敗 (${filePath}): HTTP ${code} - ${response.getContentText()}`);
    return null;
  }
}

// =================================================================
// 📤 Buffer連携用の関数群（変更なし）
// =================================================================

// デバッグ用: Bufferのトークンが有効かどうかを、投稿せずに軽くテストする
function debug_testBufferToken() {
  const url = 'https://api.buffer.com';
  const graphqlQuery = `query { user { id email } }`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ query: graphqlQuery }),
    headers: { 'Authorization': 'Bearer ' + CONFIG.ACCOUNT_1.TOKEN },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  Logger.log(`HTTPステータス: ${code}`);
  Logger.log(`レスポンス内容: ${response.getContentText()}`);
}

function postToEachChannels(token, profileIdsString, text, mediaUrl, isVideo) {
  const channelIds = profileIdsString.split(',').map(id => id.trim()).filter(id => id !== '');

  if (channelIds.length === 0) {
    // 💡 本当にIDが1つも指定されていない場合のみ、このエラーになる
    throw new Error(`有効なチャンネルIDが見つかりません。PROFILE_IDの値を確認してください: "${profileIdsString}"`);
  }

  let allSuccess = true;
  let errorMessages = [];

  channelIds.forEach((channelId) => {
    try {
      sendCreatePostMutation(token, channelId, text, mediaUrl, isVideo);
      Logger.log(`✅ チャンネル[${channelId}] への投稿に成功しました`);
      Utilities.sleep(1000);
    } catch (e) {
      allSuccess = false;
      errorMessages.push(`[${channelId}]: ${e.message}`);
      Logger.log(`❌ チャンネル[${channelId}] への投稿に失敗しました: ${e.message}`);
    }
  });

  // 💡 全チャンネルへの投稿が失敗した場合、本当のエラー内容をそのまま投げる
  if (!allSuccess) throw new Error(errorMessages.join(' | '));
  return true;
}

function sendCreatePostMutation(token, channelId, text, mediaUrl, isVideo) {
  const url = 'https://api.buffer.com';
  const graphqlQuery = `mutation CreateDirectPost($input: CreatePostInput!) { createPost(input: $input) { ... on PostActionSuccess { post { id } } ... on MutationError { message } } }`;

  const assetData = {};
  if (isVideo) {
    assetData.video = { url: mediaUrl };
  } else {
    assetData.image = { url: mediaUrl };
  }

  const payload = {
    query: graphqlQuery,
    variables: {
      input: {
        channelId: channelId,
        text: text,
        schedulingType: "automatic",
        mode: "addToQueue",
        saveToDraft: false,
        assets: [assetData]
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  };

  // 💡 502/503/504などの一時的なサーバーエラーに加え、
  //    「Address unavailable」のような接続自体の失敗も自動的に再試行する
  const maxRetries = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    let response;
    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (networkError) {
      // 接続エラー（DNS失敗、一時的な到達不能など）
      lastError = networkError;
      if (attempt < maxRetries) {
        Logger.log(`⏳ Buffer APIへの接続に失敗しました(${networkError.message})。${attempt}回目の再試行前に待機します...`);
        Utilities.sleep(3000 * attempt);
        continue;
      } else {
        throw new Error(`Buffer APIへの接続に${maxRetries}回失敗しました: ${networkError.message}`);
      }
    }

    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.errors) throw new Error(json.errors[0].message);
      const result = json.data.createPost;
      if (result && result.message) throw new Error(result.message);
      return true;
    }

    if ([502, 503, 504].includes(responseCode) && attempt < maxRetries) {
      Logger.log(`⏳ Buffer APIが一時的なエラー(HTTP ${responseCode})を返しました。${attempt}回目の再試行前に待機します...`);
      Utilities.sleep(3000 * attempt); // 3秒, 6秒と待機時間を伸ばしながら再試行
      lastError = new Error(`HTTP ${responseCode} - ${response.getContentText()}`);
      continue;
    }

    throw new Error(`HTTP ${responseCode} - ${response.getContentText()}`);
  }

  throw lastError;
}

// =================================================================
// 🔑 初回セットアップ用の関数（最初の1回だけ使います）
// =================================================================
//
// 【手順】
// 1. スクリプト プロパティに DROPBOX_APP_KEY と DROPBOX_APP_SECRET を先に登録しておく
// 2. この関数「step1_showAuthUrl」を実行し、実行ログ（表示 > 実行ログ）に出てくるURLを
//    ブラウザで開いて「許可」する
// 3. リダイレクトされた後のURL（http://localhost/?code=xxxxx の形）から
//    "code=" の後ろの文字列をコピーする
//    ※ページ自体は「読み込めません」と表示されますが、URLさえ見られればOKです
// 4. 下の「step2_exchangeCodeForRefreshToken」関数の code の中身をコピーした値に書き換えて実行
// 5. 成功すると、refresh_token が自動でスクリプト プロパティ（DROPBOX_REFRESH_TOKEN）に保存されます
// 6. この2つの関数はもう使わないので、そのまま残しておいても削除してもどちらでもOK

function step1_showAuthUrl() {
  const authUrl = 'https://www.dropbox.com/oauth2/authorize'
    + '?client_id=' + encodeURIComponent(CONFIG.DROPBOX_APP_KEY)
    + '&response_type=code'
    + '&token_access_type=offline'
    + '&redirect_uri=' + encodeURIComponent('http://localhost/');

  Logger.log('以下のURLをブラウザで開いて「許可」してください:');
  Logger.log(authUrl);
}

function step2_exchangeCodeForRefreshToken() {
  // 👇 ここを、リダイレクト先URLの code=以降の値に書き換えてください
  const code = 'ここに認証後に取得したcodeを貼り付ける';

  const url = 'https://api.dropboxapi.com/oauth2/token';
  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      code: code,
      grant_type: 'authorization_code',
      client_id: CONFIG.DROPBOX_APP_KEY,
      client_secret: CONFIG.DROPBOX_APP_SECRET,
      redirect_uri: 'http://localhost/'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode === 200) {
    const json = JSON.parse(response.getContentText());
    // 💡 ログに残さず、スクリプト プロパティへ直接保存する
    PROPS.setProperty('DROPBOX_REFRESH_TOKEN', json.refresh_token);
    Logger.log('✅ 成功！ refresh_token をスクリプト プロパティ（DROPBOX_REFRESH_TOKEN）に保存しました。');
  } else {
    Logger.log(`❌ 失敗しました: HTTP ${responseCode} - ${response.getContentText()}`);
  }
}

// =================================================================
// 🔍 デバッグ用: App Folder直下に実際にあるファイル一覧を確認する
// =================================================================
function debug_listFilesInAppFolder() {
  const accessToken = getDropboxAccessToken();
  const url = 'https://api.dropboxapi.com/2/files/list_folder';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + accessToken },
    payload: JSON.stringify({ path: '' }), // ''はApp Folderのルート
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 200) {
    const json = JSON.parse(response.getContentText());
    if (!json.entries || json.entries.length === 0) {
      Logger.log('📂 App Folder直下にファイルが1つもありません。');
    } else {
      Logger.log(`📂 App Folder直下のファイル一覧（${json.entries.length}件）:`);
      json.entries.forEach(entry => {
        Logger.log(`  - [${entry['.tag']}] "${entry.name}"  path_lower: "${entry.path_lower}"`);
      });
    }
  } else {
    Logger.log(`❌ 一覧取得に失敗しました: HTTP ${code} - ${response.getContentText()}`);
  }
}
// メニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('YouTube分析')
    .addItem('検索実行', 'searchYouTube')
    .addToUi();
}

// メイン実行関数
function searchYouTube() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('検索');
  
  // シートが存在しない場合の処理
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「検索」という名前のシートが見つかりません。作成してください。');
    return;
  }

  // 検索条件を取得
  const keyword = sheet.getRange('B2').getValue(); // 検索キーワード
  const extractionCondition = sheet.getRange('B3').getValue(); // 抽出条件
  const maxResults = sheet.getRange('B4').getValue(); // 抽出数
  const timeFilter = sheet.getRange('B5').getValue(); // 対象期間
  const sortCondition = sheet.getRange('B6').getValue(); // 並び替え基準
  const durationFilter = sheet.getRange('B7').getValue(); // 動画長　20250213追加
  const countryFilter = sheet.getRange('B8').getValue(); // 優先地域　20250218追加

  if (!keyword) {
    SpreadsheetApp.getUi().alert('検索キーワードが入力されていません。B2セルにキーワードを入力してください。');
    return;
  }

  // 抽出条件をAPIの order パラメータに変換
  let order;
  switch (extractionCondition) {
    case '関連度':
      order = 'relevance';
      break;
    case '再生回数':
      order = 'viewCount';
      break;
    case '投稿日':
      order = 'date';
      break;
    default:
      SpreadsheetApp.getUi().alert('B3セルに正しい抽出条件を選択してください。');
      return;
  }

  // 動画長をvideoDuration パラメータに変換  20250213追加
  let videoDuration;    
  switch (durationFilter) {
    case '4分未満':
      videoDuration = 'short';
      break;
    case '4分以上20分未満':
      videoDuration = 'medium';
      break;
    case '20分以上':
      videoDuration = 'long';
      break;
    case '時間指定なし':
      videoDuration = null;
      break;
    default:
      SpreadsheetApp.getUi().alert('B7セルに正しい動画長条件を選択してください');
      return;
  }

  // 優先地域フィルターをregionCodeに変換 20250218追加
  let regionCode;
  switch (countryFilter) {
    case '日本':
      regionCode = 'JP';
      break;
    case '全世界':
      regionCode = undefined;
      break;
    default:
      SpreadsheetApp.getUi().alert('B8セルに「日本」または「全世界」を選択してください。');
      return;
  }

  Logger.log(`検索キーワード: ${keyword}`);
  Logger.log(`抽出条件: ${extractionCondition} (${order})`);
  Logger.log(`最大結果数: ${maxResults}`);
  Logger.log(`対象期間: ${timeFilter}`);
  Logger.log(`並び替え条件: ${sortCondition}`);
  Logger.log(`動画長: ${videoDuration}`);
  Logger.log(`優先地域: ${regionCode}`);

  try {
    // YouTube Data APIのリクエストパラメータを定義 20250219追加
    const apiOptions = {
      q: keyword,
      maxResults: maxResults,
      order: order,
      publishedAfter: getPublishedAfterDate(timeFilter),
      type: 'video',
      videoDuration: videoDuration || undefined
    };
    if (regionCode === 'JP') {
    apiOptions.hl = 'ja'; // 日本語を優先
    apiOptions.relevanceLanguage = 'ja'; // 日本語に関連する動画を優先
    }
    if (regionCode) apiOptions.regionCode = regionCode;
  
    // YouTube Data APIのリクエスト
    const searchResults = YouTube.Search.list('snippet', apiOptions);

    if (!searchResults.items || searchResults.items.length === 0) {
      SpreadsheetApp.getUi().alert('検索結果がありません。条件を変更してください。');
      return;
    }

    Logger.log(`検索結果件数: ${searchResults.items.length}`);
    Logger.log(`APIリクエストパラメータ: ${JSON.stringify(apiOptions)}`);

    // 動画データの処理
    const videos = processVideos(searchResults.items);
    if (videos.length === 0) {
      SpreadsheetApp.getUi().alert('動画データの処理に失敗しました。');
      return;
    }

    // 並び替え適用
    const sortedVideos = sortVideos(videos, sortCondition);

    // 結果を「検索結果」シートに出力
    const outputSheet = getOrCreateSheet(ss, '検索結果');
    outputResults(sortedVideos, outputSheet);
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${error.message}`);
  }
}

// 対象期間をISO 8601形式で取得
function getPublishedAfterDate(filter) {
  const now = new Date();
  switch (filter) {
    case '今週':
      return new Date(now.setDate(now.getDate() - 7)).toISOString();
    case '今月':
      return new Date(now.setMonth(now.getMonth() - 1)).toISOString();
    case '今年':
      return new Date(now.setFullYear(now.getFullYear() - 1)).toISOString();
    default:
      return null; // 全期間
  }
}

// 動画データの処理
function processVideos(items) {
  const videoIds = items.map(item => item.id.videoId).join(',');

  // 各動画の詳細情報を取得
  const videoDetails = YouTube.Videos.list('statistics,snippet', { id: videoIds });

  return videoDetails.items.map(item => {
    const stats = item.statistics;
    const snippet = item.snippet;

    // チャンネル登録者数を取得
    const subscriberCount = getChannelSubscribers(snippet.channelId);

    return {
      title: snippet.title,
      url: `https://www.youtube.com/watch?v=${item.id}`,
      publishedAt: new Date(snippet.publishedAt),
      channelTitle: snippet.channelTitle,
      description: snippet.description,
      viewCount: parseInt(stats.viewCount) || 0,
      likeCount: parseInt(stats.likeCount) || 0,
      commentCount: parseInt(stats.commentCount) || 0,
      subscriberCount: subscriberCount
    };
  });
}

// 「グラフ」シートの全グラフを削除
function clearCharts(sheet) {
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
}

// 結果をスプレッドシートに出力
function outputResults(videos, sheet) {
  sheet.clear()

  // ヘッダー行
  const headers = ['No.', 'タイトル', 'URL', '投稿日', 'チャンネル名', '概要', 'エンゲージメント率', '急上昇率', '再生数/チャンネル登録者数'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 動画データの出力
  const outputData = videos.map((video, index) => [
    index + 1,
    video.title,
    video.url,
    video.publishedAt,
    video.channelTitle,
    video.description,
    calculateEngagementRate(video),
    calculateTrendingRate(video),
    calculateViewSubRatio(video)
  ]);

  sheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);

  // エンゲージメント率の上位5件を抽出
  const top5Videos = videos
    .map((video, index) => ({
    no: index + 1,
    title: video.title,
    engagementRate: parseFloat(calculateEngagementRate(video)),
    }))
    .sort((a, b) => b.engagementRate - a.engagementRate)
    .slice(0, 5);
  
   // 急上昇率の上位5件を抽出
  const top5TrendingVideos = videos
    .map((video, index) => ({
    no: index + 1,
    title: video.title,
    trendingRate: parseFloat(calculateTrendingRate(video)),
    }))
    .sort((a, b) => b.trendingRate - a.trendingRate)
    .slice(0, 5);

  // 再生数/チャンネル登録者数の上位5件を抽出
  const top5ViewSubRatioVideos = videos
    .map((video, index) => ({
    no: index + 1,
    title: video.title,
    viewSubRatio: parseFloat(calculateViewSubRatio(video)),
    }))
    .sort((a, b) => b.viewSubRatio - a.viewSubRatio)
    .slice(0, 5);

  // グラフ用データの出力先を取得または作成
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const graphSheet = getOrCreateSheet(ss, 'グラフ');
  clearCharts(graphSheet); // 「グラフ」シートのグラフを削除

  // 動的にグラフの開始行を調整
  let currentRow = 1; // 初期の開始行を1行目に設定

  // エンゲージメント率のグラフ
  currentRow = configureChartData(graphSheet, 'エンゲージメント率トップ5', top5Videos, 'エンゲージメント率', currentRow, 1);

  // 急上昇率のグラフ
  currentRow = configureChartData(graphSheet, '急上昇率トップ5', top5TrendingVideos, '急上昇率', currentRow, 1);

  // 再生数/チャンネル登録者数のグラフ
  configureChartData(graphSheet, '再生数/チャンネル登録者数トップ5', top5ViewSubRatioVideos, '再生数/チャンネル登録者数', currentRow, 1);

  SpreadsheetApp.getUi().alert('検索結果を「検索結果」シートに出力し、グラフを「グラフ」シートに作成しました。');
}

// エンゲージメント率の計算
function calculateEngagementRate(video) {
  return ((video.likeCount + video.commentCount) / video.viewCount * 100).toFixed(2) || 0;
}

// 急上昇率の計算
function calculateTrendingRate(video) {
  const days = (new Date() - video.publishedAt) / (1000 * 60 * 60 * 24); // 経過日数を計算
  return days > 0 ? ((video.viewCount / days) / 10000).toFixed(2) : 0; // 日数が0の場合を考慮
}

// 再生数/チャンネル登録者数
function calculateViewSubRatio(video) {
  return (video.viewCount / video.subscriberCount).toFixed(2) || 0;
}

// 期間指定
function getPublishedAfter(timeFrame) {
  const now = new Date();
  switch (timeFrame) {
    case '今週': return new Date(now.setDate(now.getDate() - 7));
    case '今月': return new Date(now.setMonth(now.getMonth() - 1));
    case '今年': return new Date(now.setFullYear(now.getFullYear() - 1));
    default: return null;
  }
}

// シートが存在しない場合は作成する
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); // 既存の場合は内容をクリア
  }
  return sheet;
}

// チャンネル登録者数を取得
function getChannelSubscribers(channelId) {
  const channelDetails = YouTube.Channels.list('statistics', { id: channelId });
  if (channelDetails.items && channelDetails.items.length > 0) {
    return parseInt(channelDetails.items[0].statistics.subscriberCount) || 0;
  }
  return 0;
}

// 動画データを並び替え
function sortVideos(videos, sortCondition) {
  switch (sortCondition) {
    case '再生回数':
      return videos.sort((a, b) => b.viewCount - a.viewCount);
    case '高評価数':
      return videos.sort((a, b) => b.likeCount - a.likeCount);
    case 'コメント数':
      return videos.sort((a, b) => b.commentCount - a.commentCount);
    case '投稿日':
      return videos.sort((a, b) => b.publishedAt - a.publishedAt);
    default:
      Logger.log(`無効な並び替え条件: ${sortCondition}`);
      return videos; // 並び替えなしで返す
  }
}

// グラフ用データを配置しグラフを生成
function configureChartData(sheet, title, data, hAxisTitle, startRow, startColumn) {
  // データの配置
  const chartHeaders = ['No.', 'タイトル', hAxisTitle];
  sheet.getRange(startRow, startColumn, 1, chartHeaders.length).setValues([chartHeaders]);
  
  // hAxisTitleに基づきデータを取得
  let mappedData;
  switch (hAxisTitle) {
    case 'エンゲージメント率':
      mappedData = data.map(item => [item.no, item.title, item.engagementRate]);
      break;
    case '急上昇率':
      mappedData = data.map(item => [item.no, item.title, item.trendingRate]);
      break;
    case '再生数/チャンネル登録者数':
      mappedData = data.map(item => [item.no, item.title, item.viewSubRatio]);
      break;
    default:
      throw new Error('不明なhAxisTitleが指定されました');
  }
  
  const dataHeight = mappedData.length + 1; // ヘッダー行 + データ行
  sheet.getRange(startRow + 1, startColumn, mappedData.length, chartHeaders.length).setValues(mappedData);

  // A列にリンクを設定
  const linkFormulas = mappedData.map(item => [
    `=HYPERLINK("#gid=${SpreadsheetApp.getActiveSpreadsheet().getSheetByName('検索結果').getSheetId()}&range=A${item[0] + 1}", "${item[0]}")`
  ]);

  // A列にリンクの数式を設定
  sheet.getRange(startRow + 1, startColumn, linkFormulas.length, 1).setFormulas(linkFormulas);

  // グラフの生成
  const chartStartRow = startRow; // グラフの上辺をデータと揃える
  createBarChart(sheet, startRow + 1, startRow + mappedData.length, title, hAxisTitle, chartStartRow, startColumn + 4);

  // 次のデータ開始行を返す（スペース込み）
  const graphHeight = 19; // グラフの高さを一定と仮定
  return chartStartRow + graphHeight + 1; // グラフに十分な高さを確保
}

 // グラフの生成
function createBarChart(sheet, startRow, endRow, title, hAxisTitle, positionRow, positionColumn) {
  const range = sheet.getRange(startRow, 2, endRow - startRow + 1, 2);
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(range)
    .setPosition(positionRow, positionColumn, 0, 0) // グラフの位置
    .setOption('title', title)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: hAxisTitle })
    .build();
  sheet.insertChart(chart);
}
function myFunction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("自動化用");
  // sheet.getRange("A1").setValue("Hello, World!");
  // const response = FetchResponseFromLLM()
  // length = response.length

  const range = sheet.getRange("C:C");
  const values = range.getValues();

  // console.log(values.length);
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === "") {
      break;
    }else{
      const response = FetchResponseFromLLM(values[i][0])
      // console.log(response)
      // var lines = response[0].split("\n");
  
      // // 各行を分割し、変数に代入
      // var item = lines[0].split("：")[1].trim();
      // var reason = lines[1].split("：")[1].trim();
      
      // console.log("項目: " + item);   // "リーダーシップ"
      // console.log("理由: " + reason);

      // sheet.getRange("C" + (i + 1)).setValue(response[0]);
      // sheet.getRange("D" + (i + 1)).setValue(response[0]);

      // console.log(extractFromJSON(response[0]));
      const [bItem, mItem, point, reason] = extractFromJSON(parseJSONSafely(response[0]))


      // 出力したいオブジェクトの準備
      const outputData = {
        "大項目": bItem,                 // 空の配列
        "中項目": mItem,                 // 空の配列
        "分類しやすさ": point,         // `分類しやすさ` には `null` を入れておく
        "理由": reason                   // `理由` は空の文字列
      };

      // console.log で出力
      console.log(bItem);

      //大項目・中項目を配列から文字列へと変換
      const [strBItem, strMItem] = formatArrayItems(bItem, mItem)

      console.log(strBItem)

      //もし大項目と中項目の数が一致しなかったら


      // sheet.getRange("H" + (i + 1)).setValue(response.join('\n'))
      sheet.getRange("G" + (i + 1)).setValue(strBItem);
      sheet.getRange("H" + (i + 1)).setValue(strMItem)
      sheet.getRange("I" + (i + 1)).setValue(reason);
      sheet.getRange("J" + (i + 1)).setValue(point)
    }
  }
  
}
  
function FetchResponseFromLLM(input) {
  //スクリプトプロパティに設定したOpenAIのAPIキーを取得
  const apiKey = ScriptProperties.getProperty('APIKEY');
  // const apiKey = PropertiesService.getScriptProperties().getProperty('APIKEY');
  //文章生成AIのAPIのエンドポイントを設定
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  //ChatGPTに投げるメッセージを定義(ユーザーロールの投稿文のみ)
  const messages = [{'role': 'user', 'content': `
  
#「上司の強み」に関するエンゲージメントサーベイを実施した。

「社員の声」：「`+input+`」
が以下の項目のどれに最も分類されるべきか教えてください。理由も考えてください。

- リーダーシップ        目標・ビジョン                
- リーダーシップ        戦略的思考                
- リーダーシップ        問題解決力                
- リーダーシップ        意思決定力                
- マネジメント        チームビルディング                
- マネジメント        業務管理                
- マネジメント        業績管理                
- コミュニケーション        フィードバック力                
- コミュニケーション        傾聴力                
- コミュニケーション        発信力                
- 人材管理        評価                
- 人材管理        育成                
- 人材管理        採用                
- 個人の資質        経験・学習                
- 個人の資質        公平性                
- 個人の資質        ストレス耐性                

また、大項目と中項目に分類する際、どのくらい分類しやすかったかを１〜５で教えて。厳しく低めの数字で教えて。
  - 1: 分類できているか分からない、間違いっている可能性が非常に高い
  - 2: 正確に分類できているか不安だ、間違っている可能性が高い
  - 3: ある程度は精度良く分類できたと思う、少し間違っている可能性がある。
  - 4: 結構しっかり分類できたと思う、あっている可能性が高い
  - 5: 正確に迷いなく分類できた、あっている可能性が非常に高い

## ステップ
1. 「`+input+`」を項目を一つずつ比較し、一番近い内容の項目を見つける。
2. その項目に分類されると思った根拠や理由を考える。
3. この判断がどのくらい正しいと思うかを1-5で評価する。

## 背景
授業員にとって働きやすい環境を提供することを目的に、まず「社員の声」を集めた。特に今回は「上司」の強みに対して意見をもらいたい。

## 役割
あなたはデータアナリストです。
社員が回答した内容を、項目別に分類したい。


## 出力形式
この中にある項目名のみ出力せよ。
「`+input+`」の内容にもっとも最も近い項目を教えて。
原則1項目まで、但し判断が難しい場合は、最大3項目まで

必ずこの中から選んでください。
- リーダーシップ        目標・ビジョン                
- リーダーシップ        戦略的思考                
- リーダーシップ        問題解決力                
- リーダーシップ        意思決定力                
- マネジメント        チームビルディング                
- マネジメント        業務管理                
- マネジメント        業績管理                
- コミュニケーション        フィードバック力                
- コミュニケーション        傾聴力                
- コミュニケーション        発信力                
- 人材管理        評価                
- 人材管理        育成                
- 人材管理        採用                
- 個人の資質        経験・学習                
- 個人の資質        公平性                
- 個人の資質        ストレス耐性             

## 出力形式
大項目、中項目、分類しやすさ（1-5）、（その項目に分類されると思った）理由
この４つの要素を含んだjsonで返して。

## 出力形式の例(json)
下記のように分類される可能性のある項目は配列で示す。
「コミュニケーション        発信力 」と「個人の資質        経験・学習 」に分類されると思った場合は次のようになる。
{
  "大項目": [リーダーシップ, コミュニケーション],
  "中項目": [問題解決力, 傾聴力],
  "分類しやすさ": 1,
  "理由": "「ノウハウの蓄積」は専門的な知識や経験を示していますが、「円滑なコミュニケーション能力」は、情報を伝え、異なる部門との連携を築く能力に関連しており、その意味で発信力に最も近いと判断しました。"
}
や
{
  "大項目": [コミュニケーション, コミュニケーション],
  "中項目": [発信力, 傾聴力],
  "分類しやすさ": 1,
  "理由": "「ノウハウの蓄積」は専門的な知識や経験を示していますが、「円滑なコミュニケーション能力」は、情報を伝え、異なる部門との連携を築く能力に関連しており、その意味で発信力に最も近いと判断しました。"
}
大項目と中項目の長さは必ず同じでなければならない。

## やってほしいこと
- 大項目、中項目、分類しやすさ、理由の４つの要素を含んだjsonで必ず出力する
- 以下の形式で必ず出力すること
{
  "大項目": [],
  "中項目": [],
  "分類しやすさ": ,
  "理由": ""
}

## やってほしくないこと
- 余計な文字を入れること（出力されたjson形式をそのまま使いたいため）

`}];

  //OpenAIのAPIリクエストに必要なヘッダー情報を設定
  const headers = {
    'Authorization':'Bearer '+ apiKey,
    'Content-type': 'application/json',
    'X-Slack-No-Retry': 1
  };
  //ChatGPTモデルやトークン上限、プロンプトをオプションに設定
  const options = {
    'muteHttpExceptions' : true,
    'headers': headers, 
    'method': 'POST',
    'payload': JSON.stringify({
      'model': 'gpt-4o-mini',
      'max_tokens' : 2048,
      'temperature' : 1,
      'messages': messages})
  };
  //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
  const response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText());
  //ChatGPTのAPIレスポンスをログ出力
  results = response.choices[0].message.content;
  // console.log(results)
  // console.log(results.length)
  // console.log(results[10])
  formattedResults = formatResponse(results)
  return formattedResults
}

function formatResponse(results){
  const resultsArray = results.split('\n\n').map(issue => issue.trim()).filter(issue => issue.length > 0);
  // console.log(issues);
  return resultsArray
}

function parseJSONSafely(jsonString) {
  try {
    // 文字列をJSONとして解析
    const parsedObject = JSON.parse(jsonString);
    return parsedObject;
  } catch (e) {
    // エラーが発生した場合、空のオブジェクトを返す
    console.error("無効なJSON文字列です。空のオブジェクトを返します。", e);
    return {}; // 空のJSONオブジェクトを返す
  }
}

function extractFromJSON(jsonData){
  console.log(jsonData)
  const bItem = jsonData['大項目']
  const mItem = jsonData['中項目']
  const point = jsonData['分類しやすさ']
  const reason = jsonData['理由']
  // console.log(bItem)

  return [bItem, mItem, point, reason]
}

function formatArrayItems(big, med) {
  return [big.join('\n'), med.join('\n')]
}


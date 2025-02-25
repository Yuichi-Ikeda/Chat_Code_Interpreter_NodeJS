const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { AzureOpenAI } = require('openai');
require('dotenv').config();

// 入力ファイルのパス
const FILE_FONT_PATH = path.join(__dirname, 'input_files', 'Font.zip');
const FILE_EXCEL_PATH = path.join(__dirname, 'input_files', 'Excel.zip');

// 環境変数から設定を取得
const apiEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
const apiKey = process.env.AZURE_OPENAI_API_KEY;
const apiVersion = process.env.API_VERSION;
const deploymentName = process.env.DEPLOYMENT_NAME;

// 非同期 IIFE で全体を実行
(async () => {
  try {
    // クライアントの初期化
    const client = new AzureOpenAI({
      apiKey: apiKey,
      apiVersion: apiVersion,
      azureEndpoint: apiEndpoint,
    });

    // FONTファイルをアップロード
    let fileStream = fs.createReadStream(FILE_FONT_PATH);
    const file_font = await client.files.create({
      file: fileStream,
      purpose: 'assistants'
    });
    fileStream.close();
    console.log(`Font file uploaded successfully. File ID: ${file_font.id}`);

    // アシスタントを作成
    const assistant = await client.beta.assistants.create({
      name: 'AI Assistant for Excel File Analysis',
      model: deploymentName,
      instructions: "You are an AI assistant that analyzes EXCEL files. Please answer user requests in Japanese.",
      tools: [{ type: 'code_interpreter' }],
      tool_resources: { code_interpreter: { file_ids: [file_font.id] } },
    });
    console.log(`Assistant created successfully. Assistant ID: ${assistant.id}`);

    // EXCEL ファイルをアップロード
    fileStream = fs.createReadStream(FILE_EXCEL_PATH);
    const file_excel = await client.files.create({
      file: fileStream,
      purpose: 'assistants'
    });
    fileStream.close();
    console.log(`Excel file uploaded successfully. File ID: ${file_excel.id}`);

    // スレッドを作成
    const thread = await client.beta.threads.create({
      messages: [{
      role: 'user',
      content: 'アップロードされた Font.zip と Excel.zip を /mnt/data/upload_files に展開してください。これらの ZIP ファイルには解析対象の EXCEL ファイルと日本語フォント NotoSansJP.ttf が含まれています。展開した先にある EXCEL ファイルをユーザーの指示に従い解析してください。EXCEL データからグラフやチャート画像を生成する場合、タイトル、軸項目、凡例等に NotoSansJP.ttf を利用してください。',
      attachments:[
        {
        "file_id": file_font.id,
        "file_id": file_excel.id,
        "tools": [{"type": "code_interpreter"}]
        }]
      }]
    });
    console.log(`Thread created successfully. Thread ID: ${thread.id}`);
    console.log("Chat session started. Type 'exit' to end the session.");

    // 標準入力の設定
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    const question = (query) =>
      new Promise((resolve) => rl.question(query, resolve));

    // チャットループ
    while (true) {
      // ユーザー入力を取得
      const user_input = await question('\nUser: ');

      // 終了コマンドの処理
      if (user_input.toLowerCase() === 'exit') {
        console.log('Ending session...');
        break;
      }

      // ユーザーのメッセージ送信
      await client.beta.threads.messages.create(
        thread.id,
        {
          role: 'user',
          content: user_input,
        }
      );

      // アシスタントの応答を取得
      let run = await client.beta.threads.runs.createAndPoll(
        thread.id,
        {
          assistant_id: assistant.id,
        },
        { pollIntervalMs: 1000 }
      );
      // console.log(`Run created:  ${JSON.stringify(run)}`);

      console.log('\nAssistant:');
      // メッセージの各コンテンツブロックを処理
      const runMessages = await client.beta.threads.messages.list(
        thread.id
      );
      for await (const runMessageDatum of runMessages) {
        for (const contentBlock of runMessageDatum.content) {
          console.log("================================");
          if (contentBlock.type === 'text') {
            console.log(contentBlock.text.value);
          } else if (contentBlock.type === 'image_file') {
            const fileId = contentBlock.image_file.file_id;
            console.log(`[Image file received: ${fileId}]`);

            try {
              // 画像データの取得
              const response = await client.files.content(fileId);
              const image_data = await response.arrayBuffer();
              const image_data_buffer = Buffer.from(image_data);

              // 画像データの保存
              fs.writeFileSync(`.\\output_images\\${fileId}.png`, image_data_buffer);
              console.log(`File saved as '${fileId}.png'`);
            } catch (error) {
              console.error(`Error retrieving image: ${error}`);
            }
          } else {
            console.log(`Unhandled content type: ${contentBlock.type}`);
          }
        }
      }
    }

    // スレッドの削除
    await client.beta.threads.del(thread.id);
    console.log('Thread deleted successfully.');

    // EXCEL ファイルを削除
    await client.files.del(file_excel.id)
    console.log("Excel file deleted successfully.")

    // アシスタントの削除
    await client.beta.assistants.del(assistant.id);
    console.log('Assistant deleted successfully.');

    // FONt ファイルを削除
    await client.files.del(file_font.id)
    console.log("Font file deleted successfully.")

    // 標準入力のクローズ
    rl.close();
  } catch (e) {
    console.error(`An error occurred: ${e}`);
  }
})();
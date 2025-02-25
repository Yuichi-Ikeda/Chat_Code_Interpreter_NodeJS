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
        attachments: [
          {
            "file_id": file_font.id,
            "file_id": file_excel.id,
            "tools": [{ "type": "code_interpreter" }]
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
      let run = await client.beta.threads.runs.create(
        thread.id,
        {
          assistant_id: assistant.id,
        }
      );
      console.log(`Run created:  ${JSON.stringify(run)}`);
      console.log(`\nWaiting for response...`);

      // チャットループ
      while (true) {
        // run の最新状態を取得
        run = await client.beta.threads.runs.retrieve(thread.id, run.id);

        // スレッド内の全メッセージを取得
        const messages = await client.beta.threads.messages.list(thread.id);

        if (run.status === 'completed') {
          console.log(`\nRun status: ${run.status}`);

          // すべてのメッセージの内容を出力（デバッグトレース用）
          messages.data.forEach(message => {
            console.log(message.content);
          });

          console.log('\nAssistant:');

          // ここでは最初のメッセージの content 配列を処理
          const contentBlocks = messages.data[0].content;
          for (const block of contentBlocks) {
            if (block.type === 'text') {
              console.log(block.text.value);
            } else if (block.type === 'image_file') {
              const fileId = block.image_file.file_id;
              console.log(`[Image file received: ${fileId}]`);
              try {
                // 元画像ファイルを取得
                const fileResponse = await client.files.content(fileId);
                const fileData = await fileResponse.buffer();
                // 保存先ディレクトリの確認と画像ファイルのローカル保存
                const outputDir = path.join(__dirname, 'output_images');
                if (!fs.existsSync(outputDir)) {
                  fs.mkdirSync(outputDir);
                }
                const filePath = path.join(outputDir, `${fileId}.png`);
                fs.writeFileSync(filePath, fileData);
                console.log(`File saved as '${fileId}.png'`);
                // 元画像ファイルを削除
                await client.files.del(fileId)
                console.log("Image file deleted successfully.")
              } catch (e) {
                console.log(`Error retrieving image: ${e.message}`);
              }
            } else {
              console.log(`Unhandled content type: ${block.type}`);
            }
          }
          break; // 内部のポーリングループを抜ける

        } else if (run.status === 'queued' || run.status === 'in_progress') {
          console.log(`\nRun status: ${run.status}`);
          // すべてのメッセージの内容を出力（デバッグトレース用）
          messages.data.forEach(message => {
            console.log(message.content);
          });
          // 5秒待機してから再度ポーリング
          await new Promise(resolve => setTimeout(resolve, 5000));

        } else {
          console.log(`Run status: ${run.status}`);
          if (run.status === 'failed') {
            console.log(`Error Code: ${run.last_error.code}, Message: ${run.last_error.message}`);
          }
          break;
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
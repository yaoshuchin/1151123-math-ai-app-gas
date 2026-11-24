const fs = require('fs');
const path = require('path');
const recordsData = JSON.parse(fs.readFileSync(path.join(__dirname, 'assets', 'records_data.json'), 'utf-8'));
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  ShadingType
} = require('docx');

function createDoc() {
  const primaryColor = "1E3A8A";   // Deep Blue
  const secondaryColor = "0D9488"; // Teal
  const darkTextColor = "1E293B";  // Slate 800
  const lightBgColor = "F8FAFC";   // Slate 50
  const calloutBorderColor = "3B82F6"; // Blue 500
  const codeBgColor = "F1F5F9";    // Slate 100

  // Helper for paragraphs
  function p(text, options = {}) {
    return new Paragraph({
      alignment: options.alignment || AlignmentType.LEFT,
      spacing: {
        before: options.before !== undefined ? options.before : 120,
        after: options.after !== undefined ? options.after : 120,
        line: 276
      },
      children: [
        new TextRun({
          text: text,
          font: "Microsoft JhengHei",
          size: options.size || 22, // 11pt
          bold: options.bold || false,
          italics: options.italics || false,
          color: options.color || darkTextColor
        })
      ]
    });
  }

  // Helper for multi-run paragraph
  function multiP(runs, options = {}) {
    return new Paragraph({
      alignment: options.alignment || AlignmentType.LEFT,
      spacing: {
        before: options.before !== undefined ? options.before : 120,
        after: options.after !== undefined ? options.after : 120,
        line: 276
      },
      children: runs.map(r => new TextRun({
        text: r.text,
        font: "Microsoft JhengHei",
        size: r.size || options.size || 22,
        bold: r.bold || false,
        italics: r.italics || false,
        color: r.color || options.color || darkTextColor
      }))
    });
  }

  // Helper for Headings
  function h1(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 360, after: 180 },
      children: [
        new TextRun({
          text: text,
          font: "Microsoft JhengHei",
          size: 32, // 16pt
          bold: true,
          color: primaryColor
        })
      ]
    });
  }

  function h2(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 120 },
      children: [
        new TextRun({
          text: text,
          font: "Microsoft JhengHei",
          size: 26, // 13pt
          bold: true,
          color: secondaryColor
        })
      ]
    });
  }

  function h3(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_3,
      spacing: { before: 180, after: 90 },
      children: [
        new TextRun({
          text: text,
          font: "Microsoft JhengHei",
          size: 24, // 12pt
          bold: true,
          color: primaryColor
        })
      ]
    });
  }

  // Helper for Callout / Quote Box
  function calloutBox(title, lines, borderColor = calloutBorderColor, bgColor = lightBgColor) {
    const tableCells = [];
    const cellChildren = [];

    if (title) {
      cellChildren.push(new Paragraph({
        spacing: { before: 80, after: 80 },
        children: [
          new TextRun({
            text: title,
            font: "Microsoft JhengHei",
            size: 22,
            bold: true,
            color: primaryColor
          })
        ]
      }));
    }

    lines.forEach(line => {
      cellChildren.push(new Paragraph({
        spacing: { before: 40, after: 40, line: 260 },
        children: [
          new TextRun({
            text: line,
            font: "Microsoft JhengHei",
            size: 20, // 10pt
            color: darkTextColor
          })
        ]
      }));
    });

    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: bgColor, type: ShadingType.CLEAR },
              margins: { top: 140, bottom: 140, left: 200, right: 200 },
              borders: {
                left: { style: BorderStyle.SINGLE, size: 24, color: borderColor },
                top: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE }
              },
              children: cellChildren
            })
          ]
        })
      ]
    });
  }

  // Helper for prompt / code box
  function promptBox(title, content) {
    const lines = content.split('\n');
    const cellChildren = [
      new Paragraph({
        spacing: { before: 60, after: 80 },
        children: [
          new TextRun({
            text: `📋 ${title}`,
            font: "Microsoft JhengHei",
            size: 22,
            bold: true,
            color: "0F766E"
          })
        ]
      })
    ];

    lines.forEach(l => {
      cellChildren.push(new Paragraph({
        spacing: { before: 30, after: 30, line: 240 },
        children: [
          new TextRun({
            text: l,
            font: "Consolas",
            size: 19,
            color: "0F172A"
          })
        ]
      }));
    });

    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: "F0FDFA", type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 180, right: 180 },
              borders: {
                left: { style: BorderStyle.SINGLE, size: 28, color: "0D9488" },
                top: { style: BorderStyle.SINGLE, size: 4, color: "CCFBF1" },
                right: { style: BorderStyle.SINGLE, size: 4, color: "CCFBF1" },
                bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCFBF1" }
              },
              children: cellChildren
            })
          ]
        })
      ]
    });
  }

  // Create table helper for slides
  function createSummaryTable(data) {
    const rows = [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: "1E3A8A", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [p("投影片編號", { bold: true, color: "FFFFFF", size: 20, alignment: AlignmentType.CENTER })]
          }),
          new TableCell({
            width: { size: 35, type: WidthType.PERCENTAGE },
            shading: { fill: "1E3A8A", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [p("標題主題", { bold: true, color: "FFFFFF", size: 20 })]
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            shading: { fill: "1E3A8A", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [p("重點說明與操作要點", { bold: true, color: "FFFFFF", size: 20 })]
          })
        ]
      })
    ];

    data.forEach((item, index) => {
      const bgColor = index % 2 === 0 ? "FFFFFF" : "F8FAFC";
      rows.push(new TableRow({
        children: [
          new TableCell({
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [p(item.slide, { size: 18, alignment: AlignmentType.CENTER, bold: true })]
          }),
          new TableCell({
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [p(item.title, { size: 18, bold: true, color: primaryColor })]
          }),
          new TableCell({
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [p(item.desc, { size: 18 })]
          })
        ]
      }));
    });

    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: rows
    });
  }

  function createRecordsTable(photos) {
    const rows = [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 12, type: WidthType.PERCENTAGE },
            shading: { fill: "BE123C", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 80, right: 80 },
            children: [p("照片編號", { bold: true, color: "FFFFFF", size: 18, alignment: AlignmentType.CENTER })]
          }),
          new TableCell({
            width: { size: 16, type: WidthType.PERCENTAGE },
            shading: { fill: "BE123C", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 80, right: 80 },
            children: [p("時間 / 分類", { bold: true, color: "FFFFFF", size: 18, alignment: AlignmentType.CENTER })]
          }),
          new TableCell({
            width: { size: 28, type: WidthType.PERCENTAGE },
            shading: { fill: "BE123C", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 80, right: 80 },
            children: [p("活動主題與情境", { bold: true, color: "FFFFFF", size: 18 })]
          }),
          new TableCell({
            width: { size: 44, type: WidthType.PERCENTAGE },
            shading: { fill: "BE123C", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 80, right: 80 },
            children: [p("現場紀實說明與教學意涵", { bold: true, color: "FFFFFF", size: 18 })]
          })
        ]
      })
    ];

    photos.forEach((item, index) => {
      const bgColor = index % 2 === 0 ? "FFFFFF" : "FFF1F2";
      rows.push(new TableRow({
        children: [
          new TableCell({
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 80, right: 80 },
            children: [p(item.id.replace('photo_', '照片 '), { bold: true, size: 18, alignment: AlignmentType.CENTER })]
          }),
          new TableCell({
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 80, right: 80 },
            children: [
              p(item.time, { bold: true, size: 17, alignment: AlignmentType.CENTER, color: "BE123C" }),
              p(`[${item.category}]`, { size: 15, alignment: AlignmentType.CENTER, color: "64748B" })
            ]
          }),
          new TableCell({
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 80, right: 80 },
            children: [
              p(item.title, { bold: true, size: 18, color: "1E293B" }),
              p(`主講/標記: ${item.speaker}`, { size: 15, color: "64748B" })
            ]
          }),
          new TableCell({
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 80, right: 80 },
            children: [p(item.description, { size: 17, color: "334155" })]
          })
        ]
      }));
    });

    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: rows
    });
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Microsoft JhengHei",
            size: 22,
            color: darkTextColor
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,    // 1 inch
              bottom: 1440,
              left: 1440,
              right: 1440
            }
          }
        },
        children: [
          // 封面大標題
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 120 },
            children: [
              new TextRun({
                text: "115 國中數學領域 數位科技推動自主學習探究",
                font: "Microsoft JhengHei",
                size: 36, // 18pt
                bold: true,
                color: primaryColor
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 200 },
            children: [
              new TextRun({
                text: "AI × APP × GAS 互動評量與雲端串接實作講義手冊",
                font: "Microsoft JhengHei",
                size: 28, // 14pt
                bold: true,
                color: secondaryColor
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 360 },
            children: [
              new TextRun({
                text: "中央課程與教學輔導諮詢教師 ｜ 桃園市數學輔導團中興國中 李慧玲老師 主講\n（完整文字辨識整理、步驟重構與教戰手冊）",
                font: "Microsoft JhengHei",
                size: 20,
                color: "64748B"
              })
            ]
          }),

          // 核心摘要導讀
          calloutBox("📘 研習核心目標與架構導覽", [
            "本教材詳細指導中學數學教師如何結合現代生成式 AI（Gemini / Canvas）與 Google 雲端生態系（Google 試算表 + Google Apps Script），以零繁複建置成本的方式，快速打造具備「關卡測驗、手寫計算紙、即時英雄榜排行、雲端成績資料庫」的數位自主學習評量 APP。",
            "全流程分為四個階段：",
            "1. 【AI 生成 APP 雛形】：透過精準教學提示詞（Prompt），在 Gemini Canvas 中生成國中數學單元互動評量網頁。",
            "2. 【試算表資料庫規劃】：建立 Google 試算表，設定檢視權限並提取 Google Sheet ID 作為資料儲存後台。",
            "3. 【GAS 部署指令生成】：向 Gemini 提供後端指令與 Sheet ID，自動生成後端 Code.gs 與前端 Index.html。",
            "4. 【Apps Script 專案發布】：將程式碼貼入 GAS、完成「所有人」存取授權、發布網頁應用程式並驗收成績回傳。"
          ]),

          p(""),
          h1("第一章：前置準備與環境設定"),

          h2("1.1 帳號選用與登入須知（簡報 Slide 2 - 3）"),
          p("進行 AI 應用開發與 Google Apps Script 雲端部署前，帳號的權限完整度至關重要。學校給予的教育帳號常因資安控管政策而鎖定外掛或自訂腳本發布功能。"),
          calloutBox("⚠️ 帳號選擇重大注意事項", [
            "✅ 建議使用：個人帳號（@gmail.com）或教育局核發之個人研習帳號（如 @ms.tyc.edu.tw）。",
            "❌ 不建議使用：學校網域給予的帳號（例如 @dyjhs.tyc.edu.tw、@mail.sjjh.tyc.edu.tw）。學校管理者常限制外部 Web App 部署與第三方授權，將導致發布失敗。"
          ], "EF4444", "FEF2F2"),
          p("登入個人 Gmail 帳號後，點擊右上角 Google 應用程式九宮格，確認已切換至正確的使用者身分，接著進入 Gemini (https://gemini.google.com/)。"),

          h2("1.2 研習協作平台 Padlet 畫面說明（簡報 Slide 4）"),
          p("研習過程中透過 Padlet 進行任務發布與成果觀摩，平台通常包含四大專區："),
          multiP([
            { text: "1. 互動式評量：", bold: true },
            { text: "提供給教師複製使用的基礎評量提示詞與任務說明。\n" },
            { text: "2. 部門 / 討論區：", bold: true },
            { text: "提供共用試算表模板與 GAS 部署提示詞指令範本。\n" },
            { text: "3. 成果_互動式評量：", bold: true },
            { text: "教師在 Gemini Canvas 生成的第一階段公開 APP 分享網址上傳區。\n" },
            { text: "4. 成果_部門：", bold: true },
            { text: "完整整合 GAS 後端與試算表後台的正式 Web App 網址上傳區。" }
          ]),

          h2("1.3 Gemini 工具設定：啟用 Canvas 模式（簡報 Slide 5）"),
          p("在進入 Gemini 對話框前，請務必先完成工具模式切換："),
          multiP([
            { text: "步驟 1：", bold: true },
            { text: "在 Gemini 輸入框旁點擊「＋」或「問問 Gemini / 多上傳選項」。\n" },
            { text: "步驟 2：", bold: true },
            { text: "在選單中點選「Canvas」（撰寫、寫程式或製作簡報）模式。\n" },
            { text: "步驟 3：", bold: true },
            { text: "確認輸入框上方或右側顯示「Canvas」已啟用。透過 Canvas 模式，AI 產生的前端網頁程式碼可直接在右側視窗中即時預覽與編輯。" }
          ]),

          p(""),
          h1("第二章：生成數學評量活動 APP"),

          h2("2.1 評量提示詞設計與解析（簡報 Slide 6）"),
          p("將預先設計的評量提示詞複製至 Gemini Canvas 輸入框。教師可依自身授課需求，自行修改「單元名稱」與「設計者姓名」。"),
          promptBox("【提示詞範例 1：生成數學評量活動 APP】", 
`你是一個專業的國中數學教師，要設計一個單元的數學評量活動，單元為：解一元一次方程式。
首頁要設計者：劉大帥老師，並有學生可以輸入：班級、座號、姓名。
題目分為三個關卡，基礎關：5題、進階題：3題、挑戰關：2題，共10題，每題10分，合計100分。
每道題目旁要有一個可以點選計算紙，提供學生計算。
三個關卡都完成後，要呈現學生在各關卡的答題情況與分數。
學生第二次重做時，題目難度一樣，但是題目要不一樣。
不要用latex語法，數學公式請在網頁中使用KaTeX(或MathJax)來渲染...`),

          p("此提示詞包含六大關鍵教學設計要素："),
          calloutBox("💡 評量活動六大教學設計規範", [
            "1. 身分識別機制：首頁必須具備班級、座號、姓名之基本登入表單。",
            "2. 階梯式關卡架構：基礎關 (5題) → 進階關 (3題) → 挑戰關 (2題)，兼顧成就感與加深加廣。",
            "3. 整合手寫計算紙：提供學生在手機或平板上隨手塗鴉演算的 Canvas 畫布，免除外找紙筆的困擾。",
            "4. 公式精緻渲染：指明使用 KaTeX 或 MathJax 渲染數學方程式，杜絕 LaTeX 原始符號外露影響學生閱讀。",
            "5. 動態題庫機制：具備第二次重作重抽題或數值變換功能，落實自主學習中「錯後再測、精熟學習」的精神。",
            "6. 關卡回饋報表：測驗結束後提供分關分數與答題明細，方便學生自主檢視學習弱點。"
          ]),

          h2("2.2 預覽成果與公開分享連結（簡報 Slide 7）"),
          p("Gemini 在右側 Canvas 視窗中會即時編譯產生網頁預覽。教師可親自輸入班級座號試玩。試測無誤後："),
          multiP([
            { text: "1. 點擊 Canvas 右上角的「分享」圖示。\n" },
            { text: "2. 系統會產生「Canvas 的公開分享連結」（例如：share.gemini.google/TpRZRFS9JH34）。\n" },
            { text: "3. 點選「複製連結」，將連結貼至 Padlet 的「成果_互動式評量」專欄，完成第一階段成果發布。" }
          ]),

          p(""),
          h1("第三章：建立雲端試算表與擷取 Sheet ID"),

          h2("3.1 建立 Google 試算表並命名（簡報 Slide 8 - 9）"),
          p("為了將所有學生的測驗結果彙整於同一處，我們利用 Google Sheets 作為雲端資料庫："),
          multiP([
            { text: "1. 開啟 Google 雲端硬碟或從應用程式九宮格進入「Google 試算表」。\n" },
            { text: "2. 點選「＋空白試算表」建立新試算表。\n" },
            { text: "3. 將左上角「未命名的試算表」修改為具有代表性的名稱（例如：「等量公理」或「一元一次方程式成績紀錄」）。" }
          ]),

          h2("3.2 試算表共用權限設定（簡報 Slide 9）"),
          p("為了確保 GAS 腳本及外部請求能順暢存取資料，必須調整試算表的分享權限："),
          calloutBox("🔑 試算表共用權限關鍵步驟", [
            "1. 點擊右上角「共用」按鈕。",
            "2. 在「一般存取權」區域，將原本的「限制」切換為「知道連結的任何人」。",
            "3. 權限角色保持「檢視者」（後端 GAS 將以擁有者身分寫入資料，前台讀取則具備開放性）。",
            "4. 點擊「複製連結」後按下「完成」。"
          ]),

          h2("3.3 擷取試算表專屬 ID (Google Sheet ID)（簡報 Slide 10）"),
          p("試算表網址中的長代碼即為其唯一的系統身分證，必須擷取出來供後續腳本使用："),
          promptBox("【網址剖析與 Sheet ID 擷取位置】",
`完整試算表網址範例：
https://docs.google.com/spreadsheets/d/1SeihVjLWtEZF4KrbtBukPqiyo4Km_0-kdcj21ee1Ex8/edit?usp=sharing

擷取規則：在 /d/ 與 /edit 之間的那一長串代碼即為 ID：
Google Sheet ID = 1SeihVjLWtEZF4KrbtBukPqiyo4Km_0-kdcj21ee1Ex8`),

          p(""),
          h1("第四章：Google Apps Script (GAS) 部署指令與架構"),

          h2("4.1 GAS 整合核心提示詞指令（簡報 Slide 11）"),
          p("取得 Sheet ID 後，回到剛剛在 Gemini Canvas 中的對話視窗。將以下指令直接貼上，並將 ID 替換為您自己的試算表 ID："),
          promptBox("【提示詞範例 2：GAS 部署與全自動化後端指令】",
`我要部署到GAS，請讓所有人繳交的檔案，整併在我的google sheet，我的google sheet ID是：
1SeihVjLWtEZF4KrbtBukPqiyo4Km_0-kdcj21ee1Ex8
請直接將此 ID 寫入程式碼中，不要預留讓我修改的空間。在我的google sheet上先預建立幾個模擬的分頁和內容。系統要能『完全自動』建立範本與模擬資料，而不需要手動進去編輯器按「執行」code.gs裡的程式碼。在網頁載入讀取資料時（getInitialData函數）加入了一個自動偵測機制。只要系統發現的Google Sheet裡面沒有相對應的分頁，請自動觸發建立並填入模擬資料。同一個座號填寫者重複測驗時，請跳出提醒視窗，『你已經有成績，請確認是否修改您的成績』。讓使用者按確定按鈕才可以讓他再次提交，但是請不要覆寫資料，直接在此座號的測驗者，下一列存儲格標示『⚠修改分數』。前台呈現即時數據時，請呈現可以辨別使用者的身份資訊，並進行排名。使用隨機代碼呈現。請不要呼叫API的模式，提供我完整的GAS程式碼，讓我可以用簡單複製貼上即可執行。部署程式碼請用 📄 符號加上命名告訴我要貼在『code.gs』前台請用 💻 加上命名明確告訴我這是貼在『Index』且統一Index的I大寫。實現CRUD功能。請處理非GAS環境中 google is not defined 的問題，加入了能讓網頁在本地預覽時提供模擬測試資料，並在部署到GAS後順暢連接到後端。同時將表單重置功能抽離為獨立函式 resetFormAfterSubmit 增加可讀性。`),

          h2("4.2 指令功能核心亮點深度解析"),
          calloutBox("🌟 GAS 指令之七大智慧架構亮點", [
            "1. 零手動初次執行（Zero-Click Initialization）：以往建立試算表表頭需手動點進 GAS 按執行，本指令在前端網頁載入呼叫 getInitialData() 時，由後端自動判斷工作表是否存在，若無則自動建立欄位與範例資料！",
            "2. 重複作答歷程保留：學生重複測驗時，彈窗提醒「你已經有成績，請確認是否修改您的成績」，確認後在次列標註『⚠修改分數』，保留歷程不強行覆寫。",
            "3. 去識別化英雄榜：為保護學生隱私，排行榜使用隨機代碼（如 T01A9K）呈現即時得分排名，兼顧競爭激勵與個資安全。",
            "4. 本地端防呆降級（Fallback）：處理在瀏覽器本地預覽時常見的 'google is not defined' 報錯，自動切換至 Mock 測試資料，部署到雲端後又無縫切回真實試算表 API。",
            "5. 明確檔案標記：輸出結果以 📄 Code.gs 與 💻 Index.html 明確分類，降低教師操作失誤。",
            "6. 模組化設計：將 resetFormAfterSubmit 獨立，增強代碼可讀性與維護性。",
            "7. 完全免 API Key：使用 Google 原生授權，不需額外申請計費 API 金鑰。"
          ]),

          h2("4.3 Gemini 指令輸出結果檢視（簡報 Slide 12 - 14, 20 - 21）"),
          p("送出指令後，Gemini 將回傳兩段核心程式碼："),
          multiP([
            { text: "• 後端程式碼 (Code.gs)：", bold: true },
            { text: "負責處理 doGet() 網頁進入點、openById() 試算表連線、getInitialData() 自動偵測建立分頁、以及 submitScore() 成績寫入。\n" },
            { text: "• 前端程式碼 (Index.html)：", bold: true },
            { text: "包含完整 React + Tailwind CSS 互動畫面、KaTeX 公式排版、手寫計算紙畫布、以及透過 google.script.run 與後端通訊的非同步機制。" }
          ]),

          p(""),
          h1("第五章：Apps Script 實作設定（Step-by-Step）"),

          h2("5.1 進入擴充功能與貼上 Code.gs（簡報 Slide 15 - 17）"),
          multiP([
            { text: "步驟 1：", bold: true },
            { text: "回到您的 Google 試算表，點擊上方選單「擴充功能」 -> 「Apps Script」。\n" },
            { text: "步驟 2：", bold: true },
            { text: "進入編輯器後，左側預設會有「程式碼.gs」（或 Code.gs）。\n" },
            { text: "步驟 3：", bold: true },
            { text: "點入後將右側編輯區原有的所有程式碼「全部清空」。\n" },
            { text: "步驟 4：", bold: true },
            { text: "回到 Gemini 複製 Code.gs 內容，完整貼入並確認 SPREADSHEET_ID 是否正確寫入。" }
          ]),

          h2("5.2 新增並貼上 Index.html（簡報 Slide 18 - 22）"),
          calloutBox("🚨 最容易出錯的關鍵細節", [
            "在 Apps Script 中新增前端 HTML 檔案時，請嚴格遵守：",
            "1. 點擊檔案列表左上方的「＋」圖示，選擇「HTML」。",
            "2. 檔案名稱務必輸入「Index」，首字母「I」一定要大寫！（若輸入小寫 index，後端 doGet() 的 HtmlService.createHtmlOutputFromFile('Index') 將找不到檔案而報錯！）。",
            "3. 建立後，將預設產生的 HTML 代碼全部清空，貼上 Gemini 產生的 Index.html 完整內容。"
          ], "EF4444", "FEF2F2"),

          h2("5.3 專案命名與儲存（簡報 Slide 23）"),
          p("1. 點選左上角「未命名的專案」，輸入專案名稱（例如：「等量公理」或「解一元一次方程式挑戰賽」）。"),
          p("2. 點擊上方工具列的「磁碟片」圖示儲存專案（或按快捷鍵 Ctrl + S）。"),
          calloutBox("⭐ 檢查清單", [
            "✔ Code.gs 已貼上並含正確試算表 ID",
            "✔ Index.html 已建立（I 大寫）且貼入完整前端代碼",
            "✔ 專案已重新命名並點選「儲存專案」",
            "滿足以上條件，即可準備進行發布部署！"
          ], "10B981", "ECFDF5"),

          p(""),
          h1("第六章：發布部署與權限授權完整流程"),

          h2("6.1 發布網頁應用程式（簡報 Slide 24 - 26）"),
          p("發布是讓學生能透過瀏覽器開啟遊戲測驗的關鍵設定。請依以下步驟操作："),
          multiP([
            { text: "1. 點擊右上角藍色「部署」按鈕，選擇「新增部署作業」。\n" },
            { text: "2. 在彈出視窗左上角點擊「齒輪圖示」（選取類型），選擇「網頁應用程式」。\n" },
            { text: "3. 填寫說明文字（例如：「第一版發布」）。\n" },
            { text: "4. 【最關鍵設定一】：執行身分 務必選擇「我 (your_email@gmail.com)」。\n" },
            { text: "5. 【最關鍵設定二】：誰可以存取 務必下拉選擇「所有人」（Anyone）。\n" },
            { text: "6. 確認兩大設定無誤後，點擊右下角「部署」按鈕。" }
          ]),
          calloutBox("⚠️ 核心設定警語", [
            "若「誰可以存取」未設定為「所有人」，學生開啟網頁時會被要求登入 Google 帳號或出現權限不足（403 錯誤），無法順利作答！"
          ], "F59E0B", "FFFBEB"),

          h2("6.2 Google 安全性警告授權四部曲（簡報 Slide 27 - 29）"),
          p("第一次部署存取試算表的應用程式時，Google 會啟動資安審查機制，請遵循以下正常授權步驟通過："),
          multiP([
            { text: "關卡一【授予存取權】：", bold: true },
            { text: "點擊彈出視窗中的「授予存取權」按鈕，並選擇您的 Google 帳戶。\n" },
            { text: "關卡二【未驗證警告】：", bold: true },
            { text: "出現「Google hasn't verified this app (Google 尚未驗證此應用程式)」紅色驚嘆號畫面。請點擊左下角的「Advanced」(進階) 展開隱藏選單。\n" },
            { text: "關卡三【前往不安全連結】：", bold: true },
            { text: "展開後點擊最下方「Go to 等量公理 (unsafe)」(前往專案名稱(不安全))。\n" },
            { text: "關卡四【確認許可】：", bold: true },
            { text: "滾動至存取權益清單最底部，點擊「Continue」或「Allow」(允許)。" }
          ]),

          h2("6.3 專案部署成功與取得網址（簡報 Slide 30）"),
          p("完成授權後，系統將顯示「專案部署成功」，並產生一串專屬的「網頁應用程式網址」："),
          promptBox("【網頁應用程式網址格式】",
`https://script.google.com/macros/s/AKfycb.../exec`),
          p("點擊「複製」按鈕儲存網址，這就是學生進行測驗與自主學習的正式入口！"),

          p(""),
          h1("第七章：驗收成果與日常維護管理"),

          h2("7.1 遊戲網址測試（簡報 Slide 31）"),
          p("在新分頁中貼上剛剛複製的 Web App 網址："),
          multiP([
            { text: "1. 確認頁面正常加載，顯示「天平守護者 等量公理大冒險」或您的單元名稱。\n" },
            { text: "2. 檢查英雄榜是否正常出現模擬資料或空白榜單。\n" },
            { text: "3. 點擊「登入開始測驗」，輸入測試班級（如 701）、座號（如 01）、姓名（如 測試員）。\n" },
            { text: "4. 體驗計算紙演算、作答題目並提交成績。" }
          ]),

          h2("7.2 試算表後台資料驗證（簡報 Slide 32）"),
          p("提交測試成績後，打開您的 Google 試算表，驗收後台成果："),
          calloutBox("📊 後台「成績紀錄」分頁結構檢驗", [
            "✔ 試算表底部自動出現名為「成績紀錄」的全新分頁。",
            "✔ 第一列標題列自動生成：【提交時間】、【班級】、【座號】、【暱稱】、【隨機代碼】、【總分】、【加法(20)】、【減法(20)】...",
            "✔ 測試填寫的成績紀錄已精確寫入第二列。",
            "✔ 重複以同座號提交時，系統跳出確認提示，確認後在次列正確標註『⚠修改分數』。"
          ]),

          h2("7.3 遊戲網址在哪裡？找回與更新部署（簡報 Slide 33 - 35）"),
          p("若關閉瀏覽器後想再次取得網址，或未來修改程式碼需要發布新版本時："),
          multiP([
            { text: "1. 打開 Google 試算表，點選「擴充功能」 -> 「Apps Script」。\n" },
            { text: "2. 點擊右上角「部署」按鈕 -> 選擇「管理部署作業」。\n" },
            { text: "3. 在「網頁應用程式」區塊即可隨時複製現有網址。\n" },
            { text: "4. 若有修改程式碼，可在管理部署作業中點選右上角「鉛筆」(編輯) 圖示，版本選擇「新版本」，再按「部署」即可更新線上版。" }
          ]),

          p(""),
          h1("第八章：全 36 頁簡報內容對照索引檢索表"),
          p("下表彙整教材投影片 1 至 36 頁之完整辨識對照表，方便教學與研習時隨時檢索："),
          createSummaryTable([
            { slide: "Slide 1", title: "首頁封面", desc: "115 國中數學領域 數位科技推動自主學習探究 - AI X APP X GAS。講師：李慧玲老師。" },
            { slide: "Slide 2", title: "登入 Gmail 帳號", desc: "建議使用個人或教育局帳號，強烈不建議使用 @學校.tyc.edu.tw 帳號。" },
            { slide: "Slide 3", title: "進入 Gemini", desc: "九宮格選 Gemini，左下角確認出現 e-mail 表示登入成功。" },
            { slide: "Slide 4", title: "Padlet 協作畫面", desc: "互動式評量、部門、成果_互動式評量（APP網址）、成果_部門（後台網址）。" },
            { slide: "Slide 5", title: "Gemini 工具設定", desc: "多上傳選項中開啟 Canvas（撰寫、寫程式或製作簡報）模式。" },
            { slide: "Slide 6", title: "1. 生成數學評量 APP", desc: "複製評量提示詞至 Gemini，修改單元名稱、設計者，產生 10 題關卡與計算紙。" },
            { slide: "Slide 7", title: "2. APP網址上傳Padlet", desc: "點 Canvas 分享圖示，複製公開連結，貼到 Padlet 成果欄位。" },
            { slide: "Slide 8", title: "3. 建立雲端試算表", desc: "從九宮格進入 Google 試算表，點選開始建立新試算表。" },
            { slide: "Slide 9", title: "4. 試算表命名與共用", desc: "改名為等量公理，共用設定切換為「知道連結的任何人: 檢視者」。" },
            { slide: "Slide 10", title: "5. 擷取 Google Sheet ID", desc: "從網址 /d/ 與 /edit 之間擷取 ID 字串。" },
            { slide: "Slide 11", title: "6. 部署GAS(1)-指令", desc: "貼上完整 GAS 後端與整合提示詞，指定自動建立工作表與防呆機制。" },
            { slide: "Slide 12", title: "6. 部署GAS(2)-下指令", desc: "回到數學遊戲 APP 對話，以 Canvas 模式送出 GAS 提示詞。" },
            { slide: "Slide 13", title: "6. 部署GAS(3)-指令結果", desc: "Gemini 產出 Code.gs（後端）與 Index.html（前端）兩份代碼。" },
            { slide: "Slide 14", title: "6. 部署GAS(4)-複製Code", desc: "點選複製內容按鈕，取得 Code.gs 後端代碼。" },
            { slide: "Slide 15", title: "6. 部署GAS(5)-進入擴充功能", desc: "在試算表選單點選「擴充功能」 -> 「Apps Script」。" },
            { slide: "Slide 16", title: "第二步 | 貼上代碼流程圖", desc: "開啟 AS -> 貼上 Code.gs -> 新增 Index.html -> 儲存專案。" },
            { slide: "Slide 17", title: "6. 部署GAS(6)-貼上Code", desc: "清空原預設內容，貼上 Code.gs 並確認包含試算表 ID。" },
            { slide: "Slide 18", title: "6. 部署GAS(7)-前端設定", desc: "點擊「＋」新增 HTML，命名輸入 Index（I 務必大寫）。" },
            { slide: "Slide 19", title: "6. 部署GAS(8)-Index畫面", desc: "進入 Index.html 編輯畫面，清空預設的 HTML 標籤。" },
            { slide: "Slide 20", title: "6. 部署GAS(9)-複製Index", desc: "回到 Gemini 複製包含 React + Tailwind CSS 的 Index.html 代碼。" },
            { slide: "Slide 21", title: "6. 部署GAS(9)-另一複製畫面", desc: "在 Frontend Application 視窗中展開代碼並點選複製內容。" },
            { slide: "Slide 22", title: "6. 部署GAS(10)-貼上Index", desc: "將代碼完整貼入 Apps Script 的 Index.html 檔案中。" },
            { slide: "Slide 23", title: "6. 部署GAS(11)-專案命名", desc: "將未命名的專案改名為等量公理，並點擊儲存按鈕。" },
            { slide: "Slide 24", title: "第三步 | 發布部署流程圖", desc: "新增部署作業 -> 選擇網頁應用程式 -> 執行身分: 我 -> 存取權: 所有人。" },
            { slide: "Slide 25", title: "6. 部署GAS(12)-專案部署", desc: "點部署 -> 新增部署作業 -> 齒輪圖示選網頁應用程式。" },
            { slide: "Slide 26", title: "6. 部署GAS(13)-存取設定", desc: "執行身分選「我」，誰可以存取選「所有人」，點選部署。" },
            { slide: "Slide 27", title: "6. 部署GAS(14)-授予存取權", desc: "彈出警告視窗，點擊授予存取權，出現未驗證畫面後點「Advanced」。" },
            { slide: "Slide 28", title: "6. 部署GAS(15)-前往不安全", desc: "點選最下方「Go to 等量公理 (unsafe)」連結。" },
            { slide: "Slide 29", title: "6. 部署GAS(16)-允許存取", desc: "閱讀權限清單，滾動至底部點擊「Continue」或「Allow」。" },
            { slide: "Slide 30", title: "6. 部署GAS(17)-部署成功", desc: "產生正式網頁應用程式網址，點選「複製」網址並按完成。" },
            { slide: "Slide 31", title: "6. 部署GAS(18)-遊戲網址測試", desc: "在新分頁開啟遊戲網址，成功進入遊戲並測試連線與英雄榜。" },
            { slide: "Slide 32", title: "6. 部署GAS(19)-確定部署成功", desc: "查看 Google 試算表，確認自動建立「成績紀錄」分頁與填入資料。" },
            { slide: "Slide 33", title: "遊戲網址在哪裡？", desc: "引導教師如何在關閉視窗後，隨時由試算表找回遊戲網址。" },
            { slide: "Slide 34", title: "6. 部署GAS(20)-管理部署", desc: "從試算表選單「擴充功能」->「Apps Script」->「部署」->「管理部署作業」。" },
            { slide: "Slide 35", title: "6. 部署GAS(21)-複製網址", desc: "在管理部署視窗複製網址，並貼至 Padlet 成果_部門欄位。" },
            { slide: "Slide 36", title: "結語與 Q&A", desc: "115 國中數學領域 數位科技推動自主學習探究 Q&A。" }
          ]),

          p(""),
          h1("第九章：Padlet 研習協作板 31 項成果與各校實作大彙整"),
          calloutBox("📌 Padlet 研習協作板資訊", [
            "協作板主題：115南投縣素養研習_Apps Script部署",
            "線上協作板短網址：https://reurl.cc/NOKz9k (https://padlet.com/wuqi/115-_apps-script-https-reurl-cc-nokz9k-s024zrw7w21c8ybhiyuy)",
            "主辦與建立者：桃園市中興國中 李慧玲老師",
            "本章節完整收錄 Padlet 5 大專欄、31 項研習成果、各校互動評量作品與備份文件連結。"
          ]),

          h2("9.1 核心教學資源與雲端文件連結"),
          multiP([
            { text: "• 📁 Code.gs 完整後端代碼雲端文件：", bold: true },
            { text: "https://docs.google.com/document/d/1wlVrxXdCPUlA8hqOVpYLicjeQSnvnzTpz80vfG2ZXas/edit?usp=sharing\n" },
            { text: "• 🧑💻 Index.html 完整前端代碼雲端文件：", bold: true },
            { text: "https://docs.google.com/document/d/1HObJlf6YgOy7evl4C2Fcm8_jQm5_3F7Bng_EDr64otk/edit?usp=sharing\n" },
            { text: "• 🎮 互動式評量示範（不只老鼠會打洞）：", bold: true },
            { text: "https://script.google.com/macros/s/AKfycbzoSwkkp3_jgehU5h8K1ilQBprDvrBXJ6DW5fyUedkFglbhCRf_LEfMtZn6WE6xfcOhCg/exec\n" },
            { text: "• 🤖 Gemini Canvas 一元一次方程式示範：", bold: true },
            { text: "https://share.gemini.google/DCEt33VRtyI2" }
          ]),

          h2("9.2 各校教師互動評量 APP 與 GAS 部署作品一覽表"),
          p("研習中來自南投與桃園各國中之現場教師，成功設計並部署之單元作品清單："),
          createSummaryTable([
            { slide: "成果 1", title: "福豐國中 ｜ 林萱老師", desc: "單元：國中數學評量【分數的四則運算】（GAS 正式網頁應用程式）" },
            { slide: "成果 2", title: "建國國中 ｜ 陳柏宇老師", desc: "單元：國中數學評量【畢氏定理】（含 Gemini Canvas 與 GAS 正式網址）" },
            { slide: "成果 3", title: "經國國中 ｜ 劉軍佑老師", desc: "單元：國中數學評量【二次函數】（含 Gemini Canvas 與 GAS 正式網址）" },
            { slide: "成果 4", title: "慈文國中 ｜ 賴佩琳老師", desc: "單元：國中數學評量【一元二次方程式闖關評量】（Q老師設計）" },
            { slide: "成果 5", title: "會稽國中 ｜ 張嘉言老師", desc: "單元：國中數學評量【解二元一次聯立方程式】（GAS 正式網頁應用程式）" },
            { slide: "成果 6", title: "福豐國中 ｜ 秘境探路人", desc: "單元：國中數學線上評量【解一元一次方程式】（GAS 正式網頁應用程式）" },
            { slide: "成果 7", title: "青溪國中 ｜ 王廷豪老師", desc: "單元：國中數學【解一元一次方程式 廷豪老師互動評量系統】" },
            { slide: "成果 8", title: "中興國中 ｜ 吳育蘋老師", desc: "單元：國中數學評量【11/16 指數律】與 AI 互動評量探索" },
            { slide: "成果 9", title: "文昌國中 ｜ 文昌國中教師", desc: "單元：學生測驗繳交與即時排名系統（GAS 部署作業）" },
            { slide: "成果 10", title: "宏仁國中 ｜ 林偉業老師", desc: "單元：國中數學自主學習評量（Gemini Canvas 成果）" },
            { slide: "成果 11", title: "魚池國中 ｜ 研習教師", desc: "單元：國中數學自主學習評量（Gemini Canvas 成果）" }
          ]),

          p(""),
          h1("第十章：南投國中現場研習教學紀實與實錄影音（24張照片與9分23秒示範影片全解析）"),
          calloutBox("📸 現場研習實況與影音專區", [
            "研習地點：南投縣立南投國中圖書館 (Dream House 自主學習空間)",
            "主講教師：教育部中央輔導團 ｜ 桃園市中興國中 李慧玲老師",
            "本章節完整記錄研習現場動態，包含李慧玲老師 9 分 23 秒之 AI Canvas 提示詞生成評量 APP 示範影片，以及 24 張活動紀實照片與教學專業說明。",
            "所有影音與照片皆已完成網頁高畫質壓縮優化，收錄於線上知識庫系統（https://yaoshuchin.github.io/1151123-math-ai-app-gas/）。"
          ]),

          h2("10.1 現場示範錄影影音全紀錄（時長 09:23）"),
          multiP([
            { text: "• 影片主題：", bold: true },
            { text: "李慧玲老師 現場投影示範：AI 提示詞全自動生成評量 APP 實況\n" },
            { text: "• 影片時長與規格：", bold: true },
            { text: "9 分 23 秒 ｜ 540p Web-Optimized H.264 (AAC 雙聲道音訊)\n" },
            { text: "• 影片重點內容：", bold: true },
            { text: "李慧玲老師在大螢幕前現場輸入精確結構化 Prompt，示範 Google Gemini Canvas 瞬時生成具備三關卡（基礎、進階、挑戰）、即時 KaTeX 數學公式排版、手寫計算紙塗鴉板與自動計分之全功能 Web APP，並現場向所有學員展示互動評量的靈活教學應用。" }
          ]),

          h2("10.2 研習現場 24 張照片紀實與教學意涵解析"),
          p("研習活動依照時間序推展，涵蓋開場引導、故事化教材、實體摺紙探究、師生即時對戰與課後共備討論："),
          createRecordsTable(recordsData.photos),

          p(""),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 360, after: 120 },
            children: [
              new TextRun({
                text: "— 本手冊由 AI 輔助系統依據研習教材深度辨識、校正與編撰成冊 —",
                font: "Microsoft JhengHei",
                size: 18,
                italics: true,
                color: "94A3B8"
              })
            ]
          })
        ]
      }
    ]
  });

  return doc;
}

async function run() {
  const doc = createDoc();
  const buffer = await Packer.toBuffer(doc);
  const outPath = path.join(__dirname, '115國中數學領域數位科技推動自主學習探究.docx');
  fs.writeFileSync(outPath, buffer);
  console.log(`Document successfully written to: ${outPath} (${buffer.length} bytes)`);
}

run().catch(err => {
  console.error("Error generating document:", err);
  process.exit(1);
});

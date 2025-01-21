// Tavily APIキーを設定
const TAVILY_API_KEY = "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";

// リスク評価ルール
const riskRules = [
  {
    type: "曖昧表現",
    pattern: /(完全に|すべて|100%|必ず|研究によると|専門家が言うには|一部では)\s*[,、.。]?/, // リスクとなる表現
    score: 2
  },
  {
    type: "感情的表現",
    pattern: /(絶対に|最高|信じられない|驚くべき|超)/, // 強い感情を伴う表現
    score: 2
  }
];

// Tavily APIを使ったファクトチェック
async function performFactCheckWithTavily(text) {
  if (!text || text.trim() === "") {
    console.warn("空のクエリを検出しました。リクエストをスキップします。");
    return { score: 0, structuralRisks: [] };
  }

  const MAX_QUERY_LENGTH = 2000;
  if (text.length > MAX_QUERY_LENGTH) {
    console.warn("クエリが大きすぎます。短縮してください。");
    throw new Error("Query size exceeds the maximum allowed length.");
  }

  try {
    const sentences = text.split(/(?<=[。！？])\s*/).filter(sentence => sentence.trim() !== "");

    for (const sentence of sentences) {
      console.log("送信するテキスト (1文):", sentence);

      const response = await fetch("https://api.tavily.com/search", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${TAVILY_API_KEY}`
        },
        body: JSON.stringify({
          query: sentence,
          search_depth: "advanced",
          include_answer: true,
          max_results: 3,
          language: "ja"
        })
      });

      if (!response.ok) {
        const errorDetails = await response.json().catch(() => null);
        console.error("APIエラーの詳細:", errorDetails);
        throw new Error(`API error: ${response.status}`);
      }

      const result = await response.json();
      console.log("API Response for sentence:", result);
      return result;
    }
  } catch (error) {
    console.error("API Error:", error);
    throw error;
  }
}

// 文法解析によるリスク評価
function detectStructuralRisks(text) {
  const riskFactors = [];

  for (const rule of riskRules) {
    if (rule.pattern && rule.pattern.test(text)) {
      const match = text.match(rule.pattern);
      if (match) {
        riskFactors.push({
          type: rule.type,
          risk: `リスク: ${rule.type} - '${match[0]}'`,
          index: match.index,
          length: match[0].length,
          score: rule.score
        });
      }
    }
  }

  return riskFactors;
}

// リスクスコアを計算
async function calculateRiskScore(text) {
  let score = 0;

  // 文法解析によるリスク評価
  const structuralRisks = detectStructuralRisks(text);
  score += structuralRisks.reduce((total, risk) => total + risk.score, 0);

  console.log(`テキスト: "${text}" に対するリスク: ${structuralRisks.map(r => r.risk).join(", ")}`);

  // APIによるリスク評価（オプション）
  try {
    const factCheckResult = await performFactCheckWithTavily(text);
    score += factCheckResult.score || 0; // APIからのスコアを加算
  } catch (error) {
    console.error("APIによるリスク評価が失敗しました:", error);
  }

  return { score, structuralRisks };
}

// ドキュメント全体の虚偽リスク分析
async function analyzeDocumentForRisks() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    let hasRisks = false;

    for (const paragraph of paragraphs.items) {
      const text = paragraph.text.trim();
      const { structuralRisks } = await calculateRiskScore(text);

      for (const risk of structuralRisks) {
        hasRisks = true;

        // リスク箇所の範囲を特定
        const range = paragraph.getRange("Whole");
        const startIndex = risk.index;
        const endIndex = startIndex + risk.length;

        // 範囲内のリスク箇所を選択してコメントを挿入
        const riskRange = range.getTextRanges([startIndex, endIndex]);
        riskRange.load("items");
        await context.sync();

        if (riskRange.items.length > 0) {
          riskRange.items[0].insertComment(risk.risk);
          console.log(`コメント挿入: ${risk.risk}`);
        }
      }
    }

    if (!hasRisks) {
      const range = body.getRange("Start");
      range.insertComment("問題ありません");
      console.log("リスクが検出されなかったため、冒頭に「問題ありません」を挿入しました。");
    }

    await context.sync();
    showMessage("ファクトチェックが完了しました。");
  });
}

// メッセージ表示
function showMessage(message, isError = false) {
  const statusDiv = document.getElementById("status");
  if (!statusDiv) {
    console.warn("ステータス表示エリアが見つかりません: #status");
    return;
  }

  statusDiv.style.display = "block";
  statusDiv.className = isError ? "error-message" : "info-message";
  statusDiv.textContent = message;
}

// イベント設定
document.addEventListener("DOMContentLoaded", () => {
  const analyzeButton = document.getElementById("analyze-document");
  if (!analyzeButton) {
    console.error("ボタンが見つかりません: #analyze-document");
    return;
  }

  analyzeButton.addEventListener("click", () => {
    console.log("ボタンがクリックされました");
    analyzeDocumentForRisks().catch((error) => {
      console.error("エラー:", error);
      showMessage("エラーが発生しました。詳細はコンソールをご確認ください。", true);
    });
  });
});

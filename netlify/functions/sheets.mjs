// Netlify Function: Google Sheets 読み書きAPI
const SHEET_ID = "17FfNTDve-vzL3n_byDV-RlncGvpQyrgdyzA5y6R29f0";

async function getAccessToken() {
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.GOOGLE_CLIENT_ID,
      client_secret: process.env.GOOGLE_CLIENT_SECRET,
      refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
      grant_type: "refresh_token",
    }),
  });
  const data = await res.json();
  if (!data.access_token) throw new Error("Token refresh failed: " + JSON.stringify(data));
  return data.access_token;
}

function parseYen(val) {
  if (!val || val.trim() === "") return 0;
  const cleaned = val.replace(/[¥,]/g, "").replace(/^-/, "MINUS").replace(/-/g, "").replace("MINUS", "-").trim();
  const num = Number(cleaned);
  return isNaN(num) ? 0 : num;
}

// 売上計画のデータ取得
async function getSalesData(token) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent("売上計画")}!A1:AX50`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json();
  const rows = data.values || [];
  if (rows.length < 2) return {};

  const months = rows[1].slice(2);

  // 対象期間: 2025/08 ~ 2026/08 (index 19~31 in months array)
  const startIdx = months.indexOf("2025/08");
  const endIdx = months.indexOf("2026/08");
  const targetMonths = months.slice(startIdx, endIdx + 1);

  // 対象行の定義 (row number 1-based)
  const targetRows = [
    { row: 3, label: "入社売上(確定)", person: "渡邉", editable: true },
    { row: 4, label: "入社売上(確定)", person: "片桐", editable: true },
    { row: 5, label: "採用支援 ※税込", person: "ダイナトレック", editable: true },
    { row: 6, label: "CA支援 ※税込", person: "Ms. Engineer", editable: true },
    { row: 11, label: "入社売上合計(確定+ヨミ)", person: "", editable: false },
    { row: 13, label: "成約金額(確定)", person: "渡邉", editable: true },
    { row: 14, label: "成約金額(確定)", person: "片桐", editable: true },
    { row: 20, label: "入金(確定)※税抜", person: "渡邉", editable: true },
    { row: 21, label: "入金(確定)※税抜", person: "片桐", editable: true },
    { row: 23, label: "入金(確定)※税込", person: "渡邉", editable: true },
    { row: 24, label: "入金(確定)※税込", person: "片桐", editable: true },
    { row: 25, label: "サーカスRA", person: "師玉", editable: true },
    { row: 26, label: "採用支援 ※税込", person: "ダイナ", editable: true },
    { row: 27, label: "CA支援 ※税込", person: "Ms. Eng", editable: true },
    { row: 28, label: "入金合計(確定)※税込", person: "", editable: false },
    { row: 29, label: "固定費", person: "", editable: true },
    { row: 30, label: "残利益", person: "", editable: false },
  ];

  const salesData = targetRows.map(({ row, label, person, editable }) => {
    const rowData = rows[row - 1] || [];
    const values = rowData.slice(2).map(parseYen);
    while (values.length < months.length) values.push(0);
    return {
      label,
      person,
      editable,
      sheetRow: row,
      values: values.slice(startIdx, endIdx + 1),
    };
  });

  return { months: targetMonths, rows: salesData, colOffset: startIdx + 2 }; // colOffset = column C start + startIdx
}

// 成約Data取得
async function getSeiyakuData(token) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent("成約Data")}!A1:U100`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json();
  const rows = data.values || [];
  if (rows.length < 2) return { headers: [], rows: [] };

  const headers = rows[1]; // Row 2 = header
  const dataRows = [];
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    // Skip empty rows (no name)
    if (!r || !r[2] || r[2].trim() === "") continue;
    dataRows.push({ sheetRow: i + 1, cells: r });
  }
  return { headers, rows: dataRows };
}

// セル更新
async function updateCells(token, updates) {
  // updates: [{ sheet, row, col, value }]
  const data = updates.map(u => {
    const colLetter = u.col < 26 ? String.fromCharCode(65 + u.col) : "A" + String.fromCharCode(65 + u.col - 26);
    return {
      range: `${u.sheet}!${colLetter}${u.row}`,
      values: [[u.value]],
    };
  });

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values:batchUpdate`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      valueInputOption: "USER_ENTERED",
      data,
    }),
  });
  return await res.json();
}

export default async (req) => {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };

  if (req.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: corsHeaders });
  }

  try {
    const token = await getAccessToken();
    const url = new URL(req.url);
    const action = url.searchParams.get("action") || "sales";

    if (req.method === "POST") {
      const body = await req.json();
      const result = await updateCells(token, body.updates);
      return new Response(JSON.stringify({ ok: true, result }), {
        status: 200,
        headers: { "Content-Type": "application/json", ...corsHeaders },
      });
    }

    let data;
    if (action === "seiyaku") {
      data = await getSeiyakuData(token);
    } else if (action === "all") {
      const [sales, seiyaku] = await Promise.all([getSalesData(token), getSeiyakuData(token)]);
      data = { sales, seiyaku };
    } else {
      data = await getSalesData(token);
    }

    return new Response(JSON.stringify(data), {
      status: 200,
      headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-cache", ...corsHeaders },
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), {
      status: 500,
      headers: { "Content-Type": "application/json", ...corsHeaders },
    });
  }
};

export const config = { path: "/api/sheets" };

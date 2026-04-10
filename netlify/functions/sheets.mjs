// Netlify Function: Google Sheets データ取得API
const SHEET_ID = "17FfNTDve-vzL3n_byDV-RlncGvpQyrgdyzA5y6R29f0";
const SHEET_NAME = "売上計画";

async function refreshToken() {
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
  if (!data.access_token) {
    throw new Error("Token refresh failed: " + JSON.stringify(data));
  }
  return data.access_token;
}

async function fetchSheetData(accessToken) {
  const encodedName = encodeURIComponent(SHEET_NAME);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodedName}!A1:AX50`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const data = await res.json();
  return data.values || [];
}

function parseYen(val) {
  if (!val || val.trim() === "") return 0;
  const cleaned = val.replace(/[¥,]/g, "").trim();
  const num = Number(cleaned);
  return isNaN(num) ? 0 : num;
}

function buildDashboardData(rows) {
  if (rows.length < 2) return {};

  const months = rows[1].slice(2);

  const data = {
    months,
    "入社売上_確定": {},
    "入社売上_ヨミ": {},
    "入社売上合計": [],
    "入金_確定_税込": {},
    "入金_ヨミ_税込": {},
    "入金合計_確定_税込": [],
    "入金合計_確定ヨミ_税込": [],
    "その他_入金": {},
  };

  for (const row of rows) {
    if (!row || row.length === 0) continue;
    const label = row[0] || "";
    const person = row[1] || "";
    let values = row.slice(2).map(parseYen);
    // Pad to match months length
    while (values.length < months.length) values.push(0);

    if (label === "入社売上(確定)") {
      data["入社売上_確定"][person] = values;
    } else if (label === "入社売上(ヨミ)") {
      data["入社売上_ヨミ"][person] = values;
    } else if (label === "入社売上合計(確定+ヨミ)") {
      data["入社売上合計"] = values;
    } else if (label === "採用支援 ※税込" && person === "ダイナトレック") {
      data["入社売上_確定"][`採用支援(${person})`] = values;
    } else if (label === "CA支援 ※税込" && person === "Ms. Engineer") {
      data["入社売上_確定"][`CA支援(${person})`] = values;
    } else if (label === "入金(確定)※税込") {
      data["入金_確定_税込"][person] = values;
    } else if (label === "入金(ヨミ)※税込") {
      data["入金_ヨミ_税込"][person] = values;
    } else if (label === "入金合計(確定)※税込") {
      data["入金合計_確定_税込"] = values;
    } else if (label === "入金合計(確定＋ヨミ)※税込") {
      data["入金合計_確定ヨミ_税込"] = values;
    } else if (label === "サーカスRA") {
      data["その他_入金"][`サーカスRA(${person})`] = values;
    } else if (label === "採用支援 ※税込" && person === "ダイナ") {
      data["その他_入金"][`採用支援(${person})`] = values;
    } else if (label === "CA支援 ※税込" && person === "Ms. Eng") {
      data["その他_入金"][`CA支援(${person})`] = values;
    } else if (label === "研修※税込") {
      data["その他_入金"][`研修(${person})`] = values;
    }
  }

  return data;
}

export default async (req) => {
  try {
    const accessToken = await refreshToken();
    const rows = await fetchSheetData(accessToken);
    const data = buildDashboardData(rows);

    return new Response(JSON.stringify(data), {
      status: 200,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "public, max-age=60",
      },
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
};

export const config = {
  path: "/api/sheets",
};

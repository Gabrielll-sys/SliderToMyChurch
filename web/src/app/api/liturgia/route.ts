import { NextRequest, NextResponse } from "next/server";

const LITURGIA_API_URL = "https://api-liturgia-diaria.vercel.app/";

function extractReference(title: string): string {
  if (!title) {
    return "";
  }
  const separator = title.indexOf(":");
  if (separator >= 0) {
    return title.slice(separator + 1).trim();
  }
  return title.trim();
}

function extractCitation(headTitle: string): string {
  if (!headTitle) {
    return "";
  }
  const match = headTitle.match(/\(([^)]+)\)/);
  if (!match?.[1]) {
    return "";
  }
  return `(${match[1].trim()})`;
}

function extractEvangelho(headTitle: string): string {
  if (!headTitle) {
    return "PROCLAMAÇÃO DO EVANGELHO";
  }
  const match = headTitle.match(/segundo\s+([^,(]+)/i);
  if (!match?.[1]) {
    return "PROCLAMAÇÃO DO EVANGELHO";
  }
  return `PROCLAMAÇÃO DO EVANGELHO DE ${match[1].trim().toUpperCase()}`;
}

export async function GET(request: NextRequest) {
  const date = request.nextUrl.searchParams.get("date");
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return NextResponse.json(
      { error: "Parâmetro date inválido. Use YYYY-MM-DD." },
      { status: 400 },
    );
  }

  try {
    const response = await fetch(`${LITURGIA_API_URL}?date=${date}`, {
      cache: "no-store",
    });

    if (!response.ok) {
      return NextResponse.json(
        { error: `Falha ao consultar API externa (status ${response.status}).` },
        { status: 502 },
      );
    }

    const data = (await response.json()) as {
      today?: {
        readings?: {
          first_reading?: { title?: string };
          second_reading?: { title?: string };
          psalm?: { title?: string; response?: string };
          gospel?: {
            head?: string;
            head_title?: string;
            head_response?: string;
          };
        };
      };
    };

    const readings = data?.today?.readings;
    if (!readings) {
      return NextResponse.json(
        { error: "Nenhum dado de liturgia encontrado para a data informada." },
        { status: 404 },
      );
    }

    const gospelTitle = readings.gospel?.head_title ?? "";

    return NextResponse.json({
      firstReadingRef: extractReference(readings.first_reading?.title ?? ""),
      psalmTitle: (readings.psalm?.title ?? "").trim(),
      psalmResponse: (readings.psalm?.response ?? "").trim(),
      secondReadingRef: extractReference(readings.second_reading?.title ?? ""),
      gospelProclamation: extractEvangelho(gospelTitle),
      gospelAcclamation: (readings.gospel?.head_response ?? "").trim(),
      gospelAntiphon: (readings.gospel?.head ?? "").trim(),
      gospelCitation: extractCitation(gospelTitle),
    });
  } catch (error) {
    return NextResponse.json(
      {
        error: `Erro ao consultar liturgia: ${
          error instanceof Error ? error.message : "desconhecido"
        }`,
      },
      { status: 500 },
    );
  }
}

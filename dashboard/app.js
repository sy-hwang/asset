let data = null;
let latest = null;
let previous = null;
let chartRecords = [];
const VALUE_GRID_STEP = 50000000;
const INVESTED_COLOR = "#ffd400";
const INVESTED_POINT_FILL = "#fffbe0";
const DEFAULT_TEXT_COLOR = "#2f3437";
const PENSION_FUND_LIMIT = 6000000;
const IRP_LIMIT = 3000000;
const ISA_OPEN_YEAR = 2024;
const ISA_ANNUAL_LIMIT = 20000000;
const ISA_MAX_LIMIT = 100000000;
const CATEGORY_PALETTE = [
  { base: "#4C9AFF", accent: "#74B2FF", tint: "#EAF3FF" },
  { base: "#2DBE7F", accent: "#57CF9A", tint: "#E9FAF2" },
  { base: "#FF6B6B", accent: "#FF8E8E", tint: "#FFEDED" },
  { base: "#FF9F43", accent: "#FFB56D", tint: "#FFF3E7" },
  { base: "#2EC4B6", accent: "#59D3C8", tint: "#E9FBF8" },
  { base: "#FFD93D", accent: "#FFE56A", tint: "#FFFBE8" },
  { base: "#A66BFF", accent: "#BC8CFF", tint: "#F2EAFF" },
  { base: "#FF7AA2", accent: "#FF9BBB", tint: "#FFF0F5" },
  { base: "#00B8D9", accent: "#33C9E4", tint: "#E8F9FC" },
  { base: "#6BCB77", accent: "#8AD894", tint: "#ECF9EE" },
  { base: "#5E81FF", accent: "#809BFF", tint: "#EDF1FF" },
  { base: "#FF8A5B", accent: "#FFA481", tint: "#FFF1EB" },
  { base: "#3DDC97", accent: "#68E4AF", tint: "#EAFBF3" },
  { base: "#9B5DE5", accent: "#B082EB", tint: "#F2EBFB" },
  { base: "#00C2A8", accent: "#33D0BA", tint: "#E7FAF6" },
  { base: "#FF5DA2", accent: "#FF82B8", tint: "#FFEAF3" },
  { base: "#38BDF8", accent: "#69CDF9", tint: "#EAF7FE" },
  { base: "#F97316", accent: "#FB9449", tint: "#FFF1E7" },
  { base: "#22C55E", accent: "#4FD17E", tint: "#E8F9EE" },
  { base: "#7C83FD", accent: "#9A9FFD", tint: "#EEF0FF" },
];

const currency = new Intl.NumberFormat("ko-KR", {
  style: "currency",
  currency: "KRW",
  maximumFractionDigits: 0,
});

const percent = new Intl.NumberFormat("ko-KR", {
  style: "percent",
  maximumFractionDigits: 1,
});

function extractTrailingCategory(text) {
  const match = String(text ?? "").match(/\(([^()]+)\)\s*$/);
  return match ? match[1].trim() : "";
}

function stripTrailingCategory(text) {
  return String(text ?? "").replace(/\s*\([^()]*\)\s*$/, "").trim();
}

function cloneData(source) {
  return JSON.parse(JSON.stringify(source));
}

function monthToComparableValue(monthText) {
  const [yearText, monthValueText] = String(monthText).split("/");
  const year = Number(yearText);
  const month = Number(monthValueText);
  return year * 100 + month;
}

function buildRecordDate(yearValue, monthValue) {
  const year = Number(yearValue);
  const month = Number(monthValue);
  if (!Number.isInteger(year) || !Number.isInteger(month)) return null;
  if (year < 2000 || year > 2100) return null;
  if (month < 1 || month > 12) return null;
  return `${year}/${String(month).padStart(2, "0")}`;
}

function sortRecordsByDate(records) {
  records.sort((a, b) => monthToComparableValue(a.date) - monthToComparableValue(b.date));
}

function sumValues(obj = {}) {
  return Object.values(obj).reduce((sum, value) => sum + Number(value || 0), 0);
}

function recalculateDerivedFields() {
  sortRecordsByDate(data.records);
  data.records.forEach((record, index) => {
    record.categories = record.categories ?? {};
    record.investedCategories = record.investedCategories ?? {};
    record.monthlyInvestedCategories = record.monthlyInvestedCategories ?? {};
    record.netWorth = sumValues(record.categories);
    record.invested = sumValues(record.investedCategories);
    record.profit = record.netWorth - record.invested;
    record.delta = index === 0 ? 0 : record.netWorth - data.records[index - 1].netWorth;
  });

  data.updatedMonth = data.records[data.records.length - 1]?.date ?? "-";
  data.latestAllocation = { ...(data.records[data.records.length - 1]?.categories ?? {}) };
}

function refreshViewState() {
  latest = data.records[data.records.length - 1];
  previous = data.records[data.records.length - 2] ?? latest;
  chartRecords = data.records.filter((item) => item.date >= "2025/09");
}

function validateDataShape(candidate) {
  return candidate && Array.isArray(candidate.records) && candidate.records.length > 0;
}

function setData(nextData) {
  if (!validateDataShape(nextData)) {
    throw new Error("유효한 데이터 형식이 아닙니다.");
  }
  data = cloneData(nextData);
  data.sourceFile = data.sourceFile ?? "Imported JSON";
  data.sheetName = data.sheetName ?? "직접 입력";
  recalculateDerivedFields();
  refreshViewState();
}

function formatKrw(value) {
  return currency.format(value);
}

function formatSigned(value) {
  const abs = formatKrw(Math.abs(value));
  if (value > 0) return `+${abs}`;
  if (value < 0) return `-${abs}`;
  return abs;
}

function formatAmountShort(value) {
  const abs = Math.abs(value);
  if (abs >= 100000000) return `${(value / 100000000).toFixed(1)}억`;
  if (abs >= 10000) return `${Math.round(value / 10000).toLocaleString("ko-KR")}만`;
  return Math.round(value).toLocaleString("ko-KR");
}

function formatEok(value) {
  const abs = Math.abs(value);
  const text = `${(abs / 100000000).toFixed(2)}억`;
  if (value > 0) return text;
  if (value < 0) return `-${text}`;
  return "0.00억";
}

function formatEokCompact(value) {
  const abs = Math.abs(value);
  const text = `${(abs / 100000000).toFixed(1)}억`;
  if (value > 0) return text;
  if (value < 0) return `-${text}`;
  return "0.0억";
}

function formatTrendMoney(value) {
  const abs = Math.abs(value);
  let text = "";

  if (abs >= 100000000) {
    text = `${(abs / 100000000).toFixed(2)}억`;
  } else if (abs >= 10000000) {
    text = `${(abs / 10000000).toFixed(1)}천`;
  } else {
    text = `${Math.round(abs / 1000000)}백`;
  }

  if (value > 0) return text;
  if (value < 0) return `-${text}`;
  return "0백";
}

function formatTrendAxisMoney(value) {
  if (Math.abs(value) < 5000000) return "0";
  return `${(value / 100000000).toFixed(1)}억`;
}

function formatManwon(value) {
  return `${Math.round((value || 0) / 10000).toLocaleString("ko-KR")}만원`;
}

function formatWonFull(value) {
  return `${Math.round(value || 0).toLocaleString("ko-KR")}원`;
}

function setText(id, value) {
  document.getElementById(id).textContent = value;
}

function makeSvgNode(tag, attrs = {}) {
  const node = document.createElementNS("http://www.w3.org/2000/svg", tag);
  Object.entries(attrs).forEach(([key, value]) => node.setAttribute(key, value));
  return node;
}

function categorizeAsset(name) {
  const category = extractTrailingCategory(name);
  if (category) return category;
  if (name.includes("보험담보대출")) return "투자";
  if (name.includes("연금") || name.includes("IRP")) return "연금";
  if (name.includes("청약") || name.includes("저축")) return "저축";
  if (name.includes("부동산") || name.includes("광교") || name.includes("보증금") || name.includes("대출")) {
    return "부동산";
  }
  return "투자";
}

function parseCategoryKey(name) {
  const category = extractTrailingCategory(name);
  if (!category) {
    return {
      itemName: String(name).trim(),
      category: categorizeAsset(String(name)),
    };
  }
  return {
    itemName: stripTrailingCategory(name),
    category,
  };
}

function getCategoryOrder() {
  const discovered = new Set();
  (data?.records ?? []).forEach((record) => {
    Object.keys(record.categories ?? {}).forEach((name) => {
      const category = categorizeAsset(name);
      if (category) discovered.add(category);
    });
    Object.keys(record.investedCategories ?? {}).forEach((name) => {
      const category = categorizeAsset(name);
      if (category) discovered.add(category);
    });
  });

  return [...discovered].sort((a, b) => a.localeCompare(b, "ko"));
}

function getCategoryPaletteIndex(category) {
  const orderedCategories = getCategoryOrder();
  const index = orderedCategories.indexOf(category);
  if (index < 0) return 0;
  return index % CATEGORY_PALETTE.length;
}

function getCategoryColor(category) {
  return CATEGORY_PALETTE[getCategoryPaletteIndex(category)].base;
}

function getCategoryAccentColor(category) {
  return CATEGORY_PALETTE[getCategoryPaletteIndex(category)].accent;
}

function getCategoryTint(category) {
  return CATEGORY_PALETTE[getCategoryPaletteIndex(category)].tint;
}

function cleanAssetName(name) {
  return name.replace(/\s*\([^)]*\)/g, "").trim();
}

function sumInvestedByKeywords(record, keywords) {
  if (!record?.investedCategories) return 0;
  const loweredKeywords = keywords.map((keyword) => String(keyword).toLowerCase());
  return Object.entries(record.investedCategories).reduce((sum, [name, value]) => {
    const loweredName = String(name).toLowerCase();
    const isMatched = loweredKeywords.some((keyword) => loweredName.includes(keyword));
    return isMatched ? sum + Number(value || 0) : sum;
  }, 0);
}

function getPreviousYearEndRecord(currentYear) {
  const candidates = data.records.filter((record) => Number(String(record.date).split("/")[0]) < currentYear);
  return candidates[candidates.length - 1] ?? null;
}

function calculateCurrentYearContribution(keywords) {
  const currentYear = Number(String(latest?.date ?? "").split("/")[0]);
  if (!Number.isFinite(currentYear)) return 0;
  const currentInvested = sumInvestedByKeywords(latest, keywords);
  const previousYearEndRecord = getPreviousYearEndRecord(currentYear);
  const previousInvested = sumInvestedByKeywords(previousYearEndRecord, keywords);
  return Math.max(currentInvested - previousInvested, 0);
}

function calculateCurrentYearMonthlyContribution(keywords, options = {}) {
  const positiveOnly = options.positiveOnly === true;
  const fromYear = Number.isFinite(options.fromYear) ? options.fromYear : null;
  const currentYear = Number(String(latest?.date ?? "").split("/")[0]);
  if (!Number.isFinite(currentYear)) return 0;
  return data.records.reduce((sum, record) => {
    const recordYear = Number(String(record.date).split("/")[0]);
    if (!Number.isFinite(recordYear)) return sum;
    if (fromYear !== null) {
      if (recordYear < fromYear || recordYear > currentYear) return sum;
    } else if (recordYear !== currentYear) {
      return sum;
    }
    const monthlyValue = sumInvestedByKeywords(
      { investedCategories: record.monthlyInvestedCategories ?? {} },
      keywords
    );
    if (positiveOnly) {
      return sum + Math.max(monthlyValue, 0);
    }
    return sum + monthlyValue;
  }, 0);
}

function getIsaLimitByYear(year) {
  const elapsedYears = Math.max(year - ISA_OPEN_YEAR + 1, 0);
  return Math.min(elapsedYears * ISA_ANNUAL_LIMIT, ISA_MAX_LIMIT);
}

function setProgressWidth(id, value, limit) {
  const node = document.getElementById(id);
  if (!node) return;
  const ratio = limit <= 0 ? 0 : Math.min(Math.max(value / limit, 0), 1);
  node.style.width = `${Math.max(ratio * 100, 0)}%`;
}

function renderSummary() {
  const changeRate = previous.netWorth === 0 ? 0 : latest.delta / previous.netWorth;
  const profitRate = latest.invested === 0 ? 0 : latest.profit / latest.invested;
  const sourceMeta = [data.sourceFile, data.sheetName].filter(Boolean).join(" · ");

  setText("updatedMonth", latest.date);
  setText("sourceFile", sourceMeta || "직접 입력 데이터");
  setText("netWorthValue", formatEok(latest.netWorth));
  setText("netWorthSubtext", `전월 대비 ${formatSigned(latest.delta)} (${percent.format(changeRate)})`);
  setText("investedValue", formatEok(latest.invested));
  setText("profitRateValue", `누적 수익률 ${percent.format(profitRate)}`);
  setText("profitValue", formatEok(latest.profit));
  setText("latestDeltaValue", `최근 한 달 증감 ${formatSigned(latest.delta)}`);
}

function renderTaxManagement() {
  const [currentYearText, currentMonthText] = String(latest?.date ?? "").split("/");
  const currentYear = Number(currentYearText);
  const currentMonth = Number(currentMonthText);
  const pensionFundPaid = calculateCurrentYearContribution(["연금저축"]);
  const irpPaid = calculateCurrentYearContribution(["irp"]);
  const isaPaidFromMonthly = calculateCurrentYearMonthlyContribution(["isa"], {
    positiveOnly: true,
    fromYear: ISA_OPEN_YEAR,
  });
  const isaPaid = isaPaidFromMonthly > 0 ? isaPaidFromMonthly : calculateCurrentYearContribution(["isa"]);
  const isaLimit = Number.isFinite(currentYear) ? getIsaLimitByYear(currentYear) : 0;
  const isaRemainingLimit = Math.max(isaLimit - isaPaid, 0);
  const remainingMonths = Number.isFinite(currentMonth) ? Math.max(13 - currentMonth, 1) : 1;
  const isaMonthlyTarget = Math.ceil(isaRemainingLimit / remainingMonths);

  setText("pensionFundPaidText", `${formatManwon(pensionFundPaid)} / ${formatManwon(PENSION_FUND_LIMIT)}`);
  setText("irpPaidText", `${formatManwon(irpPaid)} / ${formatManwon(IRP_LIMIT)}`);
  setText("isaTotalLimitText", formatWonFull(isaLimit));
  setText("isaPaidFullText", formatWonFull(isaPaid));
  setText("isaRemainingLimitText", formatWonFull(isaRemainingLimit));
  setText("isaMonthlyGuideText", `💡 매달 ${formatWonFull(isaMonthlyTarget)}씩 납입하면 올해 목표를 달성할 수 있어요!`);
  setText(
    "isaLimitMeta",
    `한도 계산 기준: ${ISA_OPEN_YEAR}년 개설, 매년 1월 ${formatManwon(ISA_ANNUAL_LIMIT)} 증액, 최대 ${formatManwon(ISA_MAX_LIMIT)}`
  );

  setProgressWidth("pensionFundProgress", pensionFundPaid, PENSION_FUND_LIMIT);
  setProgressWidth("irpProgress", irpPaid, IRP_LIMIT);
  setProgressWidth("isaProgress", isaPaid, isaLimit);
}

function shouldShowInvestedSeries(points) {
  return points.some((point) => point.invested !== null);
}

function buildLineSegments(points, valueKey) {
  const segments = [];
  let currentSegment = [];

  points.forEach((point) => {
    const value = point[valueKey];
    if (value === null || point.yMap[valueKey] === null) {
      if (currentSegment.length > 0) {
        segments.push(currentSegment);
        currentSegment = [];
      }
      return;
    }
    currentSegment.push(point);
  });

  if (currentSegment.length > 0) {
    segments.push(currentSegment);
  }

  return segments;
}

function drawSeries(svg, points, valueKey, lineClass, lineAttrs, markerAttrs) {
  const segments = buildLineSegments(points, valueKey);
  segments.forEach((segment) => {
    const d = segment
      .map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.yMap[valueKey]}`)
      .join(" ");
    svg.appendChild(makeSvgNode("path", { d, class: lineClass, ...lineAttrs }));
  });

  points.forEach((point) => {
    if (point[valueKey] === null || point.yMap[valueKey] === null) return;
    svg.appendChild(
      makeSvgNode("circle", {
        cx: point.x,
        cy: point.yMap[valueKey],
        ...markerAttrs,
      })
    );
  });
}

function drawNetWorthChart() {
  const svg = document.getElementById("netWorthChart");
  const width = 960;
  const height = 652;
  const padding = { top: 20, right: 16, bottom: 48, left: 74 };
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;
  const values = chartRecords.flatMap((item) => [item.netWorth, item.invested]).filter((value) => value !== null);
  const { min, max, range } = getSteppedTrendBounds(values, VALUE_GRID_STEP);
  const gridValues = getGridValuesByStep(min, max, VALUE_GRID_STEP);

  svg.innerHTML = "";

  gridValues.forEach((gridValue) => {
    const y = padding.top + chartHeight - ((gridValue - min) / range) * chartHeight;
    svg.appendChild(
      makeSvgNode("line", {
        x1: padding.left,
        y1: y,
        x2: width - padding.right,
        y2: y,
        class: "grid-line",
      })
    );

    const label = makeSvgNode("text", {
      x: padding.left - 14,
      y: y + 4,
      "text-anchor": "end",
      class: "tick-label",
    });
    label.textContent = `${(gridValue / 100000000).toFixed(1)}억`;
    svg.appendChild(label);
  });

  const points = chartRecords.map((item, index) => {
    const x = padding.left + (index / Math.max(chartRecords.length - 1, 1)) * chartWidth;
    return {
      ...item,
      x,
      yMap: {
        netWorth: padding.top + chartHeight - ((item.netWorth - min) / range) * chartHeight,
        invested: padding.top + chartHeight - ((item.invested - min) / range) * chartHeight,
      },
    };
  });

  if (shouldShowInvestedSeries(points)) {
    drawSeries(
      svg,
      points,
      "invested",
      "line-path-secondary",
      { stroke: INVESTED_COLOR },
      {
        r: 6,
        fill: INVESTED_POINT_FILL,
        stroke: INVESTED_COLOR,
        "stroke-width": 2.8,
      }
    );
  }

  drawSeries(
    svg,
    points,
    "netWorth",
    "line-path",
    { stroke: "#157a6e" },
    {
      r: 6.5,
      class: "point-dot",
    }
  );

  const labelStep = Math.max(Math.ceil(chartRecords.length / 6), 1);
  points.forEach((point, index) => {
    if (index % labelStep === 0 || index === chartRecords.length - 1) {
      const label = makeSvgNode("text", {
        x: point.x,
        y: height - 16,
        "text-anchor": "middle",
        class: "axis-label",
      });
      label.textContent = point.date.replace("/", ".");
      svg.appendChild(label);
    }
  });

  const hoverGuide = makeSvgNode("line", {
    x1: padding.left,
    y1: padding.top,
    x2: padding.left,
    y2: padding.top + chartHeight,
    stroke: "rgba(47, 52, 55, 0.28)",
    "stroke-width": 1,
    "stroke-dasharray": "4 4",
    visibility: "hidden",
    "pointer-events": "none",
  });
  svg.appendChild(hoverGuide);

  const hoverNetPoint = makeSvgNode("circle", {
    cx: padding.left,
    cy: padding.top,
    r: 8.2,
    fill: "#ffffff",
    stroke: "#157a6e",
    "stroke-width": 3.8,
    visibility: "hidden",
    "pointer-events": "none",
  });
  svg.appendChild(hoverNetPoint);

  const hoverInvestedPoint = makeSvgNode("circle", {
    cx: padding.left,
    cy: padding.top,
    r: 7.4,
    fill: INVESTED_POINT_FILL,
    stroke: INVESTED_COLOR,
    "stroke-width": 3.4,
    visibility: "hidden",
    "pointer-events": "none",
  });
  svg.appendChild(hoverInvestedPoint);

  const tooltipGroup = makeSvgNode("g", {
    visibility: "hidden",
    "pointer-events": "none",
  });
  const tooltipWidth = 540;
  const tooltipHeight = 276;
  const tooltipRect = makeSvgNode("rect", {
    x: 0,
    y: 0,
    width: tooltipWidth,
    height: tooltipHeight,
    rx: 18,
    fill: "rgba(255, 255, 255, 0.96)",
    stroke: "rgba(47, 52, 55, 0.14)",
  });
  const tooltipDate = makeSvgNode("text", {
    x: 28,
    y: 52,
    fill: "#2f3437",
    "font-size": "28",
    "font-weight": "700",
  });
  const tooltipNet = makeSvgNode("text", {
    x: 28,
    y: 108,
    fill: "#157a6e",
    "font-size": "24",
    "font-weight": "700",
  });
  const tooltipInvested = makeSvgNode("text", {
    x: 28,
    y: 154,
    fill: INVESTED_COLOR,
    "font-size": "24",
    "font-weight": "700",
  });
  const tooltipProfit = makeSvgNode("text", {
    x: 28,
    y: 200,
    fill: DEFAULT_TEXT_COLOR,
    "font-size": "24",
    "font-weight": "700",
  });
  const tooltipProfitRate = makeSvgNode("text", {
    x: 28,
    y: 246,
    fill: "#2f3437",
    "font-size": "24",
    "font-weight": "700",
  });
  tooltipGroup.appendChild(tooltipRect);
  tooltipGroup.appendChild(tooltipDate);
  tooltipGroup.appendChild(tooltipNet);
  tooltipGroup.appendChild(tooltipInvested);
  tooltipGroup.appendChild(tooltipProfit);
  tooltipGroup.appendChild(tooltipProfitRate);
  svg.appendChild(tooltipGroup);

  function setTooltipVisible(isVisible) {
    const visibility = isVisible ? "visible" : "hidden";
    hoverGuide.setAttribute("visibility", visibility);
    hoverNetPoint.setAttribute("visibility", visibility);
    hoverInvestedPoint.setAttribute("visibility", visibility);
    tooltipGroup.setAttribute("visibility", visibility);
  }

  function renderHover(index) {
    const point = points[index];
    if (!point) {
      setTooltipVisible(false);
      return;
    }

    setTooltipVisible(true);
    hoverGuide.setAttribute("x1", point.x);
    hoverGuide.setAttribute("x2", point.x);

    if (point.yMap.netWorth !== null) {
      hoverNetPoint.setAttribute("cx", point.x);
      hoverNetPoint.setAttribute("cy", point.yMap.netWorth);
      hoverNetPoint.setAttribute("visibility", "visible");
    } else {
      hoverNetPoint.setAttribute("visibility", "hidden");
    }

    if (point.yMap.invested !== null) {
      hoverInvestedPoint.setAttribute("cx", point.x);
      hoverInvestedPoint.setAttribute("cy", point.yMap.invested);
      hoverInvestedPoint.setAttribute("visibility", "visible");
    } else {
      hoverInvestedPoint.setAttribute("visibility", "hidden");
    }

    tooltipDate.textContent = `${point.date}`;
    tooltipNet.textContent = `순자산 ${formatKrw(point.netWorth)}`;
    tooltipInvested.textContent =
      point.invested === null ? "투자액 -" : `투자액 ${formatKrw(point.invested)}`;
    tooltipProfit.textContent =
      point.invested === null ? "수익금 -" : `수익금 ${formatKrw(point.netWorth - point.invested)}`;
    tooltipProfitRate.textContent =
      point.invested === null || point.invested <= 0
        ? "누적 수익률 -"
        : `누적 수익률 ${percent.format((point.netWorth - point.invested) / point.invested)}`;

    const tooltipX = Math.min(Math.max(point.x + 10, padding.left), width - padding.right - tooltipWidth);
    const tooltipY = Math.min(
      Math.max((point.yMap.netWorth ?? padding.top) - tooltipHeight - 8, padding.top + 2),
      padding.top + chartHeight - tooltipHeight
    );
    tooltipGroup.setAttribute("transform", `translate(${tooltipX} ${tooltipY})`);
  }

  const hoverLayer = makeSvgNode("rect", {
    x: padding.left,
    y: padding.top,
    width: chartWidth,
    height: chartHeight,
    fill: "transparent",
    style: "cursor: crosshair;",
  });
  svg.appendChild(hoverLayer);

  hoverLayer.addEventListener("mousemove", (event) => {
    const bounds = svg.getBoundingClientRect();
    const mouseX = ((event.clientX - bounds.left) / bounds.width) * width;
    const ratio = (mouseX - padding.left) / Math.max(chartWidth, 1);
    const clampedRatio = Math.min(Math.max(ratio, 0), 1);
    const index = Math.round(clampedRatio * Math.max(points.length - 1, 0));
    renderHover(index);
  });

  hoverLayer.addEventListener("mouseleave", () => {
    setTooltipVisible(false);
  });
}

function renderAllocation() {
  const donut = document.getElementById("allocationDonut");
  const list = document.getElementById("allocationList");

  const grouped = Object.entries(latest.categories).reduce((acc, [name, value]) => {
    const category = categorizeAsset(name);
    acc[category] = (acc[category] ?? 0) + value;
    return acc;
  }, {});

  const orderedEntries = getCategoryOrder()
    .map((category) => [category, grouped[category] ?? 0])
    .filter(([, value]) => value !== 0);

  const maxValue = Math.max(...orderedEntries.map(([, value]) => value), 1);
  const total = orderedEntries.reduce((sum, [, value]) => sum + value, 0);

  list.innerHTML = "";
  donut.innerHTML = "";

  const cx = 160;
  const cy = 160;
  const outerRadius = 108;
  const innerRadius = 72;
  const ringRadius = (outerRadius + innerRadius) / 2;
  const strokeWidth = outerRadius - innerRadius;
  const circumference = 2 * Math.PI * ringRadius;
  let offset = 0;

  orderedEntries.forEach(([name, value]) => {
    const color = getCategoryColor(name);
    const ratio = total === 0 ? 0 : value / total;
    const length = Math.max(ratio * circumference, 0);
    donut.appendChild(
      makeSvgNode("circle", {
        cx,
        cy,
        r: ringRadius,
        class: "allocation-slice",
        fill: "none",
        stroke: color,
        "stroke-width": strokeWidth,
        "stroke-dasharray": `${length} ${Math.max(circumference - length, 0)}`,
        "stroke-dashoffset": `${-offset}`,
        transform: `rotate(-90 ${cx} ${cy})`,
      })
    );
    offset += length;
  });

  donut.appendChild(makeSvgNode("circle", { cx, cy, r: innerRadius, fill: "#fffaf2" }));

  const centerLabel = makeSvgNode("text", {
    x: cx,
    y: cy - 8,
    class: "allocation-center-label",
  });
  centerLabel.textContent = "총 자산";
  donut.appendChild(centerLabel);

  const centerValue = makeSvgNode("text", {
    x: cx,
    y: cy + 22,
    class: "allocation-center-value",
  });
  centerValue.textContent = formatAmountShort(total);
  donut.appendChild(centerValue);

  orderedEntries.forEach(([name, value]) => {
    const color = getCategoryColor(name);
    const accentColor = getCategoryAccentColor(name);
    const item = document.createElement("article");
    item.className = "allocation-item";
    item.style.setProperty("--category-color", color);
    item.style.setProperty("--category-color-accent", accentColor);
    item.style.setProperty("--category-color-soft", `${color}1A`);

    const title = document.createElement("strong");
    title.textContent = name;

    const amount = document.createElement("span");
    amount.textContent = `${formatKrw(value)} · ${percent.format(total === 0 ? 0 : value / total)}`;

    const track = document.createElement("div");
    track.className = "allocation-track";

    const fill = document.createElement("div");
    fill.className = "allocation-fill";
    fill.style.width = `${Math.max((value / maxValue) * 100, 1)}%`;
    track.appendChild(fill);

    item.appendChild(title);
    item.appendChild(amount);
    item.appendChild(track);
    list.appendChild(item);
  });
}

function getCategoryTrend(category) {
  return data.records
    .filter((record) => record.date >= "2025/09")
    .map((record) => {
    const netWorth = Object.entries(record.categories).reduce((sum, [name, value]) => {
      return categorizeAsset(name) === category ? sum + value : sum;
    }, 0);
    const invested = Object.entries(record.investedCategories ?? {}).reduce((sum, [name, value]) => {
      return categorizeAsset(name) === category ? sum + value : sum;
    }, 0);

    return {
      date: record.date,
      netWorth,
      invested,
    };
    });
}

function getTrendValues(points) {
  return points.flatMap((item) => [item.netWorth, item.invested]).filter((value) => value !== null);
}

function getAutoTrendBounds(values) {
  if (values.length === 0) {
    return { min: 0, max: 1, range: 1 };
  }

  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const span = maxValue - minValue;
  const baseline = Math.max(Math.abs(maxValue), Math.abs(minValue), 1);
  const margin = Math.max(span * 0.15, baseline * 0.05, 1);
  const min = minValue - margin;
  const max = maxValue + margin;

  return {
    min,
    max,
    range: Math.max(max - min, 1),
  };
}

function getSteppedTrendBounds(values, step = VALUE_GRID_STEP) {
  if (values.length === 0) {
    return { min: 0, max: step, range: step };
  }
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  let min = Math.floor(minValue / step) * step;
  let max = Math.ceil(maxValue / step) * step;

  if (min === max) {
    if (min >= 0) {
      max = min + step;
      min = Math.max(0, min - step);
    } else {
      min -= step;
      max += step;
    }
  }

  return {
    min,
    max,
    range: Math.max(max - min, step),
  };
}

function getGridValuesByStep(min, max, step = VALUE_GRID_STEP) {
  const values = [];
  for (let value = min; value <= max + step * 0.001; value += step) {
    values.push(value);
  }
  return values;
}

function createScaleFromUnit(bounds, valuePerPixel, chartHeight) {
  const range = Math.max(valuePerPixel * chartHeight, bounds.range, 1);
  let min = (bounds.min + bounds.max) / 2 - range / 2;
  let max = min + range;

  // 모든 값이 양수인 경우 축을 0 아래로 내리지 않도록 보정
  if (bounds.min >= 0 && min < 0) {
    max -= min;
    min = 0;
  }

  return {
    min,
    max,
    range: Math.max(max - min, 1),
  };
}

function renderTrendChart(points, _color, height = 140, scale = null) {
  const width = 920;
  const padding = { top: 18, right: 76, bottom: 28, left: 56 };
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;
  const values = getTrendValues(points);

  if (values.length === 0) {
    return makeSvgNode("svg", { viewBox: `0 0 ${width} ${height}` });
  }

  const localBounds = getSteppedTrendBounds(values, VALUE_GRID_STEP);
  const resolvedScale =
    scale && Number.isFinite(scale.min) && Number.isFinite(scale.max)
      ? {
          min: scale.min,
          max: scale.max,
          range: Math.max(scale.max - scale.min, 1),
        }
      : scale && Number.isFinite(scale.valuePerPixel)
        ? createScaleFromUnit(localBounds, scale.valuePerPixel, chartHeight)
        : localBounds;
  const { min, max, range } = resolvedScale;

  const svg = makeSvgNode("svg", {
    viewBox: `0 0 ${width} ${height}`,
    role: "img",
    "aria-label": "월별 추이 그래프",
  });

  const gridValues = getGridValuesByStep(min, max, VALUE_GRID_STEP);
  gridValues.forEach((gridValue) => {
    const y = padding.top + chartHeight - ((gridValue - min) / range) * chartHeight;
    svg.appendChild(
      makeSvgNode("line", {
        x1: padding.left,
        y1: y,
        x2: width - padding.right,
        y2: y,
        class: "detail-trend-grid",
      })
    );

    const axisLabel = makeSvgNode("text", {
      x: 0,
      y: y + 8,
      "text-anchor": "start",
      class: "detail-trend-label",
    });
    axisLabel.textContent = formatTrendAxisMoney(gridValue);
    svg.appendChild(axisLabel);
  });

  const mapped = points.map((item, index) => {
    const x = padding.left + (index / Math.max(points.length - 1, 1)) * chartWidth;
    return {
      ...item,
      x,
      yMap: {
        netWorth:
          item.netWorth === null ? null : padding.top + chartHeight - ((item.netWorth - min) / range) * chartHeight,
        invested:
          item.invested === null ? null : padding.top + chartHeight - ((item.invested - min) / range) * chartHeight,
      },
    };
  });

  if (shouldShowInvestedSeries(mapped)) {
    drawSeries(
      svg,
      mapped,
      "invested",
      "line-path-secondary",
      { stroke: INVESTED_COLOR },
      {
        r: 6.6,
        fill: INVESTED_POINT_FILL,
        stroke: INVESTED_COLOR,
        "stroke-width": 3.2,
      }
    );
  }

  drawSeries(
    svg,
    mapped,
    "netWorth",
    "detail-trend-line",
    { stroke: "#157a6e" },
    {
      r: 7.2,
      class: "detail-trend-point",
      fill: "#ffffff",
      stroke: "#157a6e",
    }
  );

  const first = mapped[0];
  const last = mapped[mapped.length - 1];

  const labelIndexes = [...new Set([0, 2, 4, mapped.length - 1].filter((index) => index >= 0 && index < mapped.length))];
  labelIndexes.forEach((index) => {
    const point = mapped[index];
    const label = makeSvgNode("text", {
      x: point.x,
      y: height - 8,
      "text-anchor": index === 0 ? "start" : index === mapped.length - 1 ? "end" : "middle",
      class: "detail-trend-label",
    });
    label.textContent = point.date.replace("/", ".");
    svg.appendChild(label);
  });

  const hoverGuide = makeSvgNode("line", {
    x1: padding.left,
    y1: padding.top,
    x2: padding.left,
    y2: padding.top + chartHeight,
    stroke: "rgba(47, 52, 55, 0.28)",
    "stroke-width": 1,
    "stroke-dasharray": "4 4",
    visibility: "hidden",
    "pointer-events": "none",
  });
  svg.appendChild(hoverGuide);

  const hoverNetPoint = makeSvgNode("circle", {
    cx: padding.left,
    cy: padding.top,
    r: 8.2,
    fill: "#ffffff",
    stroke: "#157a6e",
    "stroke-width": 3.8,
    visibility: "hidden",
    "pointer-events": "none",
  });
  svg.appendChild(hoverNetPoint);

  const hoverInvestedPoint = makeSvgNode("circle", {
    cx: padding.left,
    cy: padding.top,
    r: 7.4,
    fill: INVESTED_POINT_FILL,
    stroke: INVESTED_COLOR,
    "stroke-width": 3.4,
    visibility: "hidden",
    "pointer-events": "none",
  });
  svg.appendChild(hoverInvestedPoint);

  const tooltipGroup = makeSvgNode("g", {
    visibility: "hidden",
    "pointer-events": "none",
  });
  const tooltipWidth = 540;
  const tooltipHeight = 276;
  const tooltipRect = makeSvgNode("rect", {
    x: 0,
    y: 0,
    width: tooltipWidth,
    height: tooltipHeight,
    rx: 18,
    fill: "rgba(255, 255, 255, 0.96)",
    stroke: "rgba(47, 52, 55, 0.14)",
  });
  const tooltipDate = makeSvgNode("text", {
    x: 28,
    y: 52,
    fill: "#2f3437",
    "font-size": "28",
    "font-weight": "700",
  });
  const tooltipNet = makeSvgNode("text", {
    x: 28,
    y: 108,
    fill: "#157a6e",
    "font-size": "24",
    "font-weight": "700",
  });
  const tooltipInvested = makeSvgNode("text", {
    x: 28,
    y: 154,
    fill: INVESTED_COLOR,
    "font-size": "24",
    "font-weight": "700",
  });
  const tooltipProfit = makeSvgNode("text", {
    x: 28,
    y: 200,
    fill: DEFAULT_TEXT_COLOR,
    "font-size": "24",
    "font-weight": "700",
  });
  const tooltipProfitRate = makeSvgNode("text", {
    x: 28,
    y: 246,
    fill: "#2f3437",
    "font-size": "24",
    "font-weight": "700",
  });
  tooltipGroup.appendChild(tooltipRect);
  tooltipGroup.appendChild(tooltipDate);
  tooltipGroup.appendChild(tooltipNet);
  tooltipGroup.appendChild(tooltipInvested);
  tooltipGroup.appendChild(tooltipProfit);
  tooltipGroup.appendChild(tooltipProfitRate);
  svg.appendChild(tooltipGroup);

  function setTooltipVisible(isVisible) {
    const visibility = isVisible ? "visible" : "hidden";
    hoverGuide.setAttribute("visibility", visibility);
    hoverNetPoint.setAttribute("visibility", visibility);
    hoverInvestedPoint.setAttribute("visibility", visibility);
    tooltipGroup.setAttribute("visibility", visibility);
  }

  function renderHover(index) {
    const point = mapped[index];
    if (!point) {
      setTooltipVisible(false);
      return;
    }

    setTooltipVisible(true);
    hoverGuide.setAttribute("x1", point.x);
    hoverGuide.setAttribute("x2", point.x);

    if (point.yMap.netWorth !== null) {
      hoverNetPoint.setAttribute("cx", point.x);
      hoverNetPoint.setAttribute("cy", point.yMap.netWorth);
      hoverNetPoint.setAttribute("visibility", "visible");
    } else {
      hoverNetPoint.setAttribute("visibility", "hidden");
    }

    if (point.yMap.invested !== null) {
      hoverInvestedPoint.setAttribute("cx", point.x);
      hoverInvestedPoint.setAttribute("cy", point.yMap.invested);
      hoverInvestedPoint.setAttribute("visibility", "visible");
    } else {
      hoverInvestedPoint.setAttribute("visibility", "hidden");
    }

    tooltipDate.textContent = `${point.date}`;
    tooltipNet.textContent = `순자산 ${formatKrw(point.netWorth)}`;
    tooltipInvested.textContent =
      point.invested === null ? "투자액 -" : `투자액 ${formatKrw(point.invested)}`;
    tooltipProfit.textContent =
      point.invested === null ? "수익금 -" : `수익금 ${formatKrw(point.netWorth - point.invested)}`;
    tooltipProfitRate.textContent =
      point.invested === null || point.invested <= 0
        ? "누적 수익률 -"
        : `누적 수익률 ${percent.format((point.netWorth - point.invested) / point.invested)}`;

    const tooltipX = Math.min(Math.max(point.x + 10, padding.left), width - padding.right - tooltipWidth);
    const tooltipY = Math.min(
      Math.max((point.yMap.netWorth ?? padding.top) - tooltipHeight - 8, padding.top + 2),
      padding.top + chartHeight - tooltipHeight
    );
    tooltipGroup.setAttribute("transform", `translate(${tooltipX} ${tooltipY})`);
  }

  const hoverLayer = makeSvgNode("rect", {
    x: padding.left,
    y: padding.top,
    width: chartWidth,
    height: chartHeight,
    fill: "transparent",
    style: "cursor: crosshair;",
  });
  svg.appendChild(hoverLayer);

  hoverLayer.addEventListener("mousemove", (event) => {
    const bounds = svg.getBoundingClientRect();
    const mouseX = ((event.clientX - bounds.left) / bounds.width) * width;
    const ratio = (mouseX - padding.left) / Math.max(chartWidth, 1);
    const clampedRatio = Math.min(Math.max(ratio, 0), 1);
    const index = Math.round(clampedRatio * Math.max(mapped.length - 1, 0));
    renderHover(index);
  });

  hoverLayer.addEventListener("mouseleave", () => {
    setTooltipVisible(false);
  });

  return svg;
}

function renderDetailCards() {
  const container = document.getElementById("detailCards");
  const coreTrendPlotHeight = 652 - 20 - 48;
  const detailTrendPlotHeight = coreTrendPlotHeight / 3;
  const detailTrendHeight = detailTrendPlotHeight + 18 + 28;
  const grouped = Object.entries(latest.categories).reduce((acc, [name, value]) => {
    const category = categorizeAsset(name);
    if (!acc[category]) acc[category] = [];
    acc[category].push({ name, value });
    return acc;
  }, {});

  container.innerHTML = "";

  const visibleCategories = getCategoryOrder();
  const trendSeriesMap = Object.fromEntries(visibleCategories.map((category) => [category, getCategoryTrend(category)]));

  getCategoryOrder().forEach((category) => {
    const items = [...(grouped[category] ?? [])].sort((a, b) => {
      if (category === "부동산") {
        const aNegative = a.value < 0 ? 1 : 0;
        const bNegative = b.value < 0 ? 1 : 0;
        if (aNegative !== bNegative) return aNegative - bNegative;
      }
      return Math.abs(b.value) - Math.abs(a.value);
    });

    if (items.length === 0) return;

    const total = items.reduce((sum, item) => sum + item.value, 0);
    const maxAbs = Math.max(...items.map((item) => Math.abs(item.value)), 1);
    const color = getCategoryColor(category);
    const chartHeight = detailTrendHeight;

    const card = document.createElement("article");
    card.className = "detail-card";
    card.style.background = getCategoryTint(category);

    const header = document.createElement("div");
    header.className = "detail-card-header";

    const headerText = document.createElement("div");
    const title = document.createElement("h3");
    title.className = "detail-card-title";
    title.textContent = category;

    const totalText = document.createElement("p");
    totalText.className = "detail-card-total";
    totalText.textContent = formatTrendMoney(total);

    headerText.appendChild(title);
    headerText.appendChild(totalText);

    const chip = document.createElement("span");
    chip.className = "detail-chip";
    chip.style.background = color;

    header.appendChild(headerText);
    header.appendChild(chip);
    card.appendChild(header);

    const body = document.createElement("div");
    body.className = "detail-card-body";

    const trendWrap = document.createElement("div");
    trendWrap.className = "detail-trend";
    trendWrap.appendChild(renderTrendChart(trendSeriesMap[category], color, chartHeight, null));
    body.appendChild(trendWrap);

    const list = document.createElement("div");
    list.className = "detail-list";

    items.forEach((item) => {
      const isPositiveTotal = item.value > 0;
      const investedValue = Number(latest.investedCategories?.[item.name] ?? 0);
      const shouldShowInvestedOverlay = Math.abs(item.value - investedValue) > 0.5;
      const row = document.createElement("article");
      row.className = "detail-item";

      const name = document.createElement("strong");
      name.textContent = cleanAssetName(item.name);

      const amount = document.createElement("span");
      amount.textContent = formatKrw(item.value);
      amount.style.color = item.value < 0 ? "#e06a52" : "#2c9b62";
      amount.style.fontWeight = "700";

      const bar = document.createElement("div");
      bar.className = "detail-bar";

      if (isPositiveTotal) {
        const totalFill = document.createElement("div");
        totalFill.className = "detail-bar-fill detail-bar-fill-total";
        totalFill.style.width = `${Math.max((Math.abs(item.value) / maxAbs) * 100, 4)}%`;
        bar.appendChild(totalFill);

        if (shouldShowInvestedOverlay) {
          const investedFill = document.createElement("div");
          investedFill.className = "detail-bar-fill detail-bar-fill-invested";
          investedFill.style.width = `${Math.max((Math.abs(investedValue) / maxAbs) * 100, 0)}%`;
          bar.appendChild(investedFill);
        }
      } else {
        const fill = document.createElement("div");
        fill.className = "detail-bar-fill detail-bar-fill-negative";
        fill.style.width = `${Math.max((Math.abs(item.value) / maxAbs) * 100, 4)}%`;
        bar.appendChild(fill);
      }

      row.appendChild(name);
      row.appendChild(amount);
      row.appendChild(bar);
      list.appendChild(row);
    });

    body.appendChild(list);
    card.appendChild(body);
    container.appendChild(card);
  });
}

function renderDashboard() {
  renderSummary();
  drawNetWorthChart();
  renderAllocation();
  renderDetailCards();
  renderTaxManagement();
}

function showFeedback(message, isError = false) {
  const feedback = document.getElementById("entryFeedback");
  if (!feedback) return;
  feedback.textContent = message;
  feedback.style.color = isError ? "#c7392f" : "#1f7a66";
}

function normalizeHeaderName(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()_\-]/g, "");
}

function resolveHeaderKey(headers, candidates) {
  const normalizedCandidates = candidates.map((item) => normalizeHeaderName(item));
  return headers.find((header) => normalizedCandidates.includes(normalizeHeaderName(header))) ?? null;
}

function toRecordDateFromExcelValue(value) {
  if (
    typeof value === "number" &&
    Number.isFinite(value) &&
    // Excel serial date 허용 범위만 날짜로 해석 (금액 숫자 오인 방지)
    value >= 20000 &&
    value <= 80000 &&
    window.XLSX?.SSF?.parse_date_code
  ) {
    const parsedDate = window.XLSX.SSF.parse_date_code(value);
    if (parsedDate?.y && parsedDate?.m) {
      return buildRecordDate(parsedDate.y, parsedDate.m);
    }
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return buildRecordDate(value.getFullYear(), value.getMonth() + 1);
  }
  const text = String(value ?? "").trim();
  if (!text) return null;

  // 엑셀 시리얼이 문자열로 들어온 경우 (예: "45292")
  if (/^\d{5}$/.test(text)) {
    const serial = Number(text);
    if (serial >= 20000 && serial <= 80000 && window.XLSX?.SSF?.parse_date_code) {
      const parsedDate = window.XLSX.SSF.parse_date_code(serial);
      if (parsedDate?.y && parsedDate?.m) {
        return buildRecordDate(parsedDate.y, parsedDate.m);
      }
    }
  }

  // YYYY/MM, YYYY-MM, YYYY.MM, YYYY년 M월 등 구분자 기반 포맷
  const separatedMatch = text.match(/^(\d{4})[\/\-.년\s]+(0?[1-9]|1[0-2])(?:월)?$/);
  if (separatedMatch) {
    return buildRecordDate(separatedMatch[1], separatedMatch[2]);
  }

  // 텍스트 앞뒤에 부가 문구가 있는 경우 (예: "기준 2024.08", "2024년 8월 말")
  const separatedInText = text.match(/(\d{4})[\/\-.년\s]+(0?[1-9]|1[0-2])(?:월)?/);
  if (separatedInText) {
    return buildRecordDate(separatedInText[1], separatedInText[2]);
  }

  // YY/MM, YY-MM, YY.MM 포맷 (예: 23.01 -> 2023/01)
  const shortYearMatch = text.match(/^(\d{2})[\/\-.년\s]+(0?[1-9]|1[0-2])(?:월)?$/);
  if (shortYearMatch) {
    const year = 2000 + Number(shortYearMatch[1]);
    return buildRecordDate(year, shortYearMatch[2]);
  }

  // 2자리 연도 + 부가 문구 허용 (예: "24.08 월", "기준 24-8")
  const shortYearInText = text.match(/(\d{2})[\/\-.년\s]+(0?[1-9]|1[0-2])(?:월)?/);
  if (shortYearInText) {
    const year = 2000 + Number(shortYearInText[1]);
    return buildRecordDate(year, shortYearInText[2]);
  }

  // YYYYMM 포맷
  const compactMatch = text.match(/^(\d{4})(0[1-9]|1[0-2])$/);
  if (compactMatch) {
    return buildRecordDate(compactMatch[1], compactMatch[2]);
  }

  return null;
}

function parseNumberValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const rawText = String(value ?? "").trim();
  if (!rawText) return 0;

  // (1,234) 형태의 음수 표기 지원
  const isParenNegative = /^\(.*\)$/.test(rawText);
  const stripped = isParenNegative ? rawText.slice(1, -1) : rawText;

  // 통화/한글/공백/기타 기호 제거 후 숫자/부호/소수점만 유지
  const normalized = stripped
    .replace(/,/g, "")
    .replace(/[₩원\s]/g, "")
    .replace(/[^0-9.\-]/g, "");

  if (!normalized || normalized === "-" || normalized === ".") return 0;
  const numericBase = Number(normalized);
  const numeric = isParenNegative ? -numericBase : numericBase;
  return Number.isFinite(numeric) ? numeric : 0;
}

function extractCategoryFromItem(itemName, fallbackCategory = "") {
  const category = extractTrailingCategory(itemName);
  if (category) return category;
  const fallback = String(fallbackCategory ?? "").trim();
  return fallback || "투자";
}

function buildDataFromExcelRows(rows, sourceFileName, sheetName) {
  if (!rows.length) {
    throw new Error("엑셀 시트에 데이터가 없습니다.");
  }

  const headers = Object.keys(rows[0]);
  const dateKey = resolveHeaderKey(headers, ["date", "월", "기준월", "month"]);
  const itemKey = resolveHeaderKey(headers, ["item", "세부항목", "세부항목명", "항목명"]);
  const categoryKey = resolveHeaderKey(headers, ["category", "카테고리", "분류"]);
  const netWorthKey = resolveHeaderKey(headers, ["networth", "평가금액", "순자산", "자산"]);
  const investedKey = resolveHeaderKey(headers, ["invested", "투자원금", "원금", "투자금"]);
  const monthlyInvestedKey = resolveHeaderKey(headers, ["월투자액", "월투자", "당월투자", "월납입", "monthlyinvested"]);

  if (!dateKey || !itemKey || !categoryKey || !netWorthKey || !investedKey) {
    throw new Error("엑셀 헤더를 찾을 수 없습니다. (월/세부항목명/카테고리/평가금액/투자원금 필요)");
  }

  const recordMap = new Map();
  rows.forEach((row) => {
    const date = toRecordDateFromExcelValue(row[dateKey]);
    const itemName = String(row[itemKey] ?? "").trim();
    const category = String(row[categoryKey] ?? "").trim();
    if (!date || !itemName || !category) return;

    if (!recordMap.has(date)) {
      recordMap.set(date, {
        date,
        categories: {},
        investedCategories: {},
        monthlyInvestedCategories: {},
      });
    }

    const record = recordMap.get(date);
    const key = `${itemName} (${category})`;
    record.categories[key] = parseNumberValue(row[netWorthKey]);
    record.investedCategories[key] = parseNumberValue(row[investedKey]);
    if (monthlyInvestedKey) {
      record.monthlyInvestedCategories[key] = parseNumberValue(row[monthlyInvestedKey]);
    }
  });

  const records = [...recordMap.values()];
  if (!records.length) {
    throw new Error("유효한 행을 찾지 못했습니다. 값이 비어있는지 확인해주세요.");
  }

  return {
    sourceFile: sourceFileName || "Imported.xlsx",
    sheetName: sheetName || "Sheet1",
    updatedMonth: "",
    records,
    latestAllocation: {},
  };
}

function findHeaderRowIndex(matrix) {
  function getMonthCountForRow(row = []) {
    let yearContext = null;
    let count = 0;
    row.forEach((cell) => {
      const date = toRecordDateFromExcelValue(cell);
      if (date) {
        yearContext = extractYearFromRecordDate(date) ?? yearContext;
        count += 1;
        return;
      }
      const monthOnly = parseMonthOnlyLabel(cell);
      if (monthOnly !== null && Number.isFinite(yearContext)) {
        count += 1;
      }
    });
    return count;
  }

  let bestIndex = -1;
  let bestScore = 0;
  matrix.forEach((row, index) => {
    const score = getMonthCountForRow(row);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });
  if (bestScore >= 3) return bestIndex;

  return -1;
}

function getParsedDataDiagnostics(parsed) {
  const records = [...(parsed?.records ?? [])];
  sortRecordsByDate(records);
  const latest = records[records.length - 1] ?? null;
  const latestYear = latest ? Number(String(latest.date).split("/")[0]) : null;
  const nonZeroMonthCount = records.filter((record) => sumValues(record.categories ?? {}) !== 0).length;
  return {
    recordCount: records.length,
    latestDate: latest?.date ?? "-",
    latestYear: Number.isFinite(latestYear) ? latestYear : null,
    latestNetWorth: latest ? sumValues(latest.categories ?? {}) : 0,
    nonZeroMonthCount,
  };
}

function inferMetricTypeFromText(text) {
  const normalized = String(text ?? "").replace(/\s+/g, "");
  if (!normalized) return null;
  if (isDerivedMetricText(normalized)) {
    return null;
  }
  const isMonthlyFlowMetric =
    normalized.includes("월간") ||
    normalized.includes("월투자") ||
    normalized.includes("당월") ||
    normalized.includes("이번달") ||
    normalized.includes("증분") ||
    normalized.includes("추가") ||
    normalized.includes("입금") ||
    normalized.includes("출금");

  const isInvestedMetric =
    normalized.includes("투자원금") ||
    normalized.includes("원금") ||
    normalized.includes("매입금액") ||
    normalized.includes("납입금액") ||
    normalized.includes("투자액") ||
    normalized.includes("누적투자") ||
    normalized.includes("누적원금") ||
    normalized.includes("총투자") ||
    normalized.includes("invested") ||
    normalized.includes("월투자");

  if (isMonthlyFlowMetric && isInvestedMetric) {
    return "monthlyInvested";
  }

  if (
    isInvestedMetric &&
    !isMonthlyFlowMetric
  ) {
    return "invested";
  }
  if (
    normalized.includes("평가금액") ||
    normalized.includes("순자산") ||
    normalized.includes("총액") ||
    normalized.includes("평가액") ||
    normalized.includes("잔액") ||
    normalized.includes("순자산액") ||
    normalized.includes("자산총액") ||
    normalized.includes("networth")
  ) {
    return "netWorth";
  }
  return null;
}

function isDerivedMetricText(text) {
  const normalized = String(text ?? "").replace(/\s+/g, "");
  if (!normalized) return false;
  return (
    normalized.includes("손익") ||
    normalized.includes("수익") ||
    normalized.includes("증감") ||
    normalized.includes("변동")
  );
}

function extractYearFromRecordDate(recordDate) {
  const match = String(recordDate ?? "").match(/^(\d{4})\/(0[1-9]|1[0-2])$/);
  return match ? Number(match[1]) : null;
}

function parseMonthOnlyLabel(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const monthMatch = text.match(/^(0?[1-9]|1[0-2])(?:월)?$/);
  if (!monthMatch) return null;
  return Number(monthMatch[1]);
}

function buildDataFromExcelMatrix(matrix, sourceFileName, sheetName) {
  const headerRowIndex = findHeaderRowIndex(matrix);
  if (headerRowIndex < 0) {
    throw new Error("월 헤더 행을 찾지 못했습니다.");
  }

  const headerRow = matrix[headerRowIndex] ?? [];
  const monthColumns = [];
  let yearContext = null;
  headerRow.forEach((cell, colIndex) => {
    const date = toRecordDateFromExcelValue(cell);
    if (date) {
      yearContext = extractYearFromRecordDate(date) ?? yearContext;
      monthColumns.push({ colIndex, date });
      return;
    }

    // "2월", "03" 같은 월-only 헤더를 같은 행의 연도 컨텍스트로 보정
    const monthOnly = parseMonthOnlyLabel(cell);
    if (monthOnly !== null && Number.isFinite(yearContext)) {
      const dateFromMonthOnly = `${yearContext}/${String(monthOnly).padStart(2, "0")}`;
      monthColumns.push({ colIndex, date: dateFromMonthOnly });
    }
  });

  if (monthColumns.length === 0) {
    throw new Error("월 헤더를 찾지 못했습니다.");
  }

  const recordMap = new Map();
  let currentMetricType = "netWorth";

  for (let rowIndex = headerRowIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const firstCell = String(row[0] ?? "").trim();
    const secondCell = String(row[1] ?? "").trim();
    const sectionText = [firstCell, secondCell].filter(Boolean).join(" ");
    const sectionMetric = inferMetricTypeFromText(sectionText);
    if (sectionMetric) {
      currentMetricType = sectionMetric;
      continue;
    }

    const itemNameRaw = firstCell || secondCell;
    if (!itemNameRaw) continue;

    const categoryFromCol = secondCell;
    const category = extractCategoryFromItem(itemNameRaw, categoryFromCol);
    const pureItemName = stripTrailingCategory(itemNameRaw);
    if (!pureItemName) continue;
    const key = `${pureItemName} (${category})`;

    let hasAnyValue = false;
    monthColumns.forEach(({ colIndex, date }) => {
      const numeric = parseNumberValue(row[colIndex]);
      if (!recordMap.has(date)) {
        recordMap.set(date, {
          date,
          categories: {},
          investedCategories: {},
          monthlyInvestedCategories: {},
        });
      }
      const target = recordMap.get(date);
      if (currentMetricType === "monthlyInvested") {
        target.monthlyInvestedCategories[key] = numeric;
      } else if (currentMetricType === "invested") {
        target.investedCategories[key] = numeric;
      } else {
        target.categories[key] = numeric;
      }
      if (numeric !== 0) {
        hasAnyValue = true;
      }
    });

    if (!hasAnyValue) {
      continue;
    }
  }

  const records = [...recordMap.values()];
  if (!records.length) {
    throw new Error("유효한 시트 데이터(월/항목/금액)를 찾지 못했습니다.");
  }

  records.forEach((record) => {
    record.categories = record.categories ?? {};
    record.investedCategories = record.investedCategories ?? {};
    record.monthlyInvestedCategories = record.monthlyInvestedCategories ?? {};
    const allKeys = new Set([
      ...Object.keys(record.categories),
      ...Object.keys(record.investedCategories),
      ...Object.keys(record.monthlyInvestedCategories),
    ]);
    allKeys.forEach((key) => {
      if (!(key in record.categories)) record.categories[key] = 0;
      if (!(key in record.investedCategories)) record.investedCategories[key] = 0;
      if (!(key in record.monthlyInvestedCategories)) record.monthlyInvestedCategories[key] = 0;
    });
  });

  return {
    sourceFile: sourceFileName || "Imported.xlsx",
    sheetName: sheetName || "Sheet1",
    updatedMonth: "",
    records,
    latestAllocation: {},
  };
}

function findDateColumnInfo(matrix) {
  let bestCol = -1;
  let bestCount = 0;
  let bestFirstRow = -1;
  const maxCols = Math.max(...matrix.map((row) => row.length), 0);

  for (let col = 0; col < maxCols; col += 1) {
    let count = 0;
    let firstRow = -1;
    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
      const date = toRecordDateFromExcelValue(matrix[rowIndex]?.[col]);
      if (!date) continue;
      if (firstRow < 0) firstRow = rowIndex;
      count += 1;
    }
    if (count > bestCount) {
      bestCount = count;
      bestCol = col;
      bestFirstRow = firstRow;
    }
  }

  if (bestCol < 0 || bestCount < 3) {
    return null;
  }

  return { dateColIndex: bestCol, dateRowStart: bestFirstRow, dateRowCount: bestCount };
}

function resolveColumnItemName(matrix, colIndex, dateRowStart) {
  // 규칙: 2행(인덱스 1)은 세부자산명
  const secondRowItem = getMergedAwareHeaderText(matrix, 1, colIndex);
  if (
    secondRowItem &&
    !inferMetricTypeFromText(secondRowItem) &&
    !isDerivedMetricText(secondRowItem) &&
    !toRecordDateFromExcelValue(secondRowItem)
  ) {
    return secondRowItem;
  }

  for (let row = dateRowStart - 1; row >= 0; row -= 1) {
    const value = getMergedAwareHeaderText(matrix, row, colIndex);
    if (!value) continue;
    // row3 같은 지표명은 항목명으로 사용하지 않고, 상위 자산명(예: ISA, IRP)을 찾는다.
    if (inferMetricTypeFromText(value) || isDerivedMetricText(value)) {
      continue;
    }
    return value;
  }
  return "";
}

function resolveColumnMetricType(matrix, colIndex, dateRowStart) {
  for (let row = dateRowStart - 1; row >= 0; row -= 1) {
    const value = getMergedAwareHeaderText(matrix, row, colIndex);
    const metric = inferMetricTypeFromText(value);
    if (metric) return metric;
  }
  return null;
}

function resolveColumnCategory(matrix, colIndex, dateRowStart) {
  // 규칙: 1행(인덱스 0)은 카테고리
  const firstRowCategory = getMergedAwareHeaderText(matrix, 0, colIndex);
  if (
    firstRowCategory &&
    !inferMetricTypeFromText(firstRowCategory) &&
    !isDerivedMetricText(firstRowCategory) &&
    !toRecordDateFromExcelValue(firstRowCategory)
  ) {
    return firstRowCategory;
  }

  for (let row = dateRowStart - 1; row >= 0; row -= 1) {
    const value = getMergedAwareHeaderText(matrix, row, colIndex);
    if (!value) continue;
    if (inferMetricTypeFromText(value) || isDerivedMetricText(value)) continue;
    if (toRecordDateFromExcelValue(value)) continue;
    return value;
  }
  return "";
}

function getMergedAwareHeaderText(matrix, rowIndex, colIndex) {
  const row = matrix[rowIndex] ?? [];
  const direct = String(row[colIndex] ?? "").trim();
  if (direct) return direct;

  // 엑셀 병합 셀은 좌상단 셀만 값이 있고 나머지는 빈칸이므로, 왼쪽으로 역탐색해 대표값을 가져온다.
  for (let left = colIndex - 1; left >= 0; left -= 1) {
    const candidate = String(row[left] ?? "").trim();
    if (!candidate) continue;
    return candidate;
  }

  return "";
}

function buildDataFromExcelDateRows(matrix, sourceFileName, sheetName) {
  const dateInfo = findDateColumnInfo(matrix);
  if (!dateInfo) {
    throw new Error("A열 월-행 구조를 찾지 못했습니다.");
  }

  const { dateColIndex, dateRowStart } = dateInfo;
  const maxCols = Math.max(...matrix.map((row) => row.length), 0);
  const columnMappings = [];

  for (let col = 0; col < maxCols; col += 1) {
    if (col === dateColIndex) continue;
    const itemName = resolveColumnItemName(matrix, col, dateRowStart);
    if (!itemName) continue;
    const metricType = resolveColumnMetricType(matrix, col, dateRowStart);
    if (!metricType) continue;
    const categoryFromHeader = resolveColumnCategory(matrix, col, dateRowStart);
    const category = extractCategoryFromItem(itemName, categoryFromHeader);
    const pureItemName = stripTrailingCategory(itemName);
    if (!pureItemName) continue;
    const key = `${pureItemName} (${category})`;
    columnMappings.push({ colIndex: col, key, metricType });
  }

  if (columnMappings.length === 0) {
    throw new Error("자산 컬럼 매핑을 만들지 못했습니다.");
  }

  // 같은 자산 키에 동일 metricType 컬럼이 여러 개면 마지막 컬럼 하나만 사용 (중복 덮어쓰기 방지)
  const uniqueMappings = [];
  const mappingIndexByKey = new Map();
  columnMappings.forEach((mapping) => {
    const mapKey = `${mapping.key}::${mapping.metricType}`;
    if (mappingIndexByKey.has(mapKey)) {
      uniqueMappings[mappingIndexByKey.get(mapKey)] = mapping;
    } else {
      mappingIndexByKey.set(mapKey, uniqueMappings.length);
      uniqueMappings.push(mapping);
    }
  });

  const keyHasNetWorthColumn = new Set(
    uniqueMappings.filter((mapping) => mapping.metricType === "netWorth").map((mapping) => mapping.key)
  );

  const recordMap = new Map();
  for (let rowIndex = dateRowStart; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const date = toRecordDateFromExcelValue(row[dateColIndex]);
    if (!date) continue;

    if (!recordMap.has(date)) {
      recordMap.set(date, {
        date,
        categories: {},
        investedCategories: {},
        monthlyInvestedCategories: {},
      });
    }

    const target = recordMap.get(date);
    uniqueMappings.forEach(({ colIndex, key, metricType }) => {
      const numeric = parseNumberValue(row[colIndex]);
      if (metricType === "monthlyInvested") {
        target.monthlyInvestedCategories[key] = numeric;
      } else if (metricType === "invested") {
        target.investedCategories[key] = numeric;
        // 일부 항목(예: 보증금/대출)은 총액 컬럼 없이 투자액만 존재한다.
        // 이 경우 투자액을 순자산 값으로도 사용해 누락을 방지한다.
        if (!keyHasNetWorthColumn.has(key)) {
          target.categories[key] = numeric;
        }
      } else {
        target.categories[key] = numeric;
      }
    });
  }

  const records = [...recordMap.values()];
  if (!records.length) {
    throw new Error("월별 레코드를 찾지 못했습니다.");
  }

  records.forEach((record) => {
    record.monthlyInvestedCategories = record.monthlyInvestedCategories ?? {};
    const allKeys = new Set([
      ...Object.keys(record.categories),
      ...Object.keys(record.investedCategories),
      ...Object.keys(record.monthlyInvestedCategories),
    ]);
    allKeys.forEach((key) => {
      if (!(key in record.categories)) record.categories[key] = 0;
      if (!(key in record.investedCategories)) record.investedCategories[key] = 0;
      if (!(key in record.monthlyInvestedCategories)) record.monthlyInvestedCategories[key] = 0;
    });
  });

  return {
    sourceFile: sourceFileName || "Imported.xlsx",
    sheetName: sheetName || "Sheet1",
    updatedMonth: "",
    records,
    latestAllocation: {},
  };
}

function parseWorkbookToData(workbook, sourceFileName) {
  if (!workbook?.SheetNames?.length) {
    throw new Error("엑셀 워크북이 비어 있습니다.");
  }

  const parseErrors = [];
  const debugCandidates = [];
  let bestCandidate = null;

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const rows = window.XLSX.utils.sheet_to_json(sheet, {
      defval: "",
    });
    const matrix = window.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: true,
    });

    try {
      const parsed = buildDataFromExcelRows(rows, sourceFileName, sheetName);
      const diagnostics = getParsedDataDiagnostics(parsed);
      debugCandidates.push({
        sheetName,
        parser: "rows",
        ...diagnostics,
      });
      const isRecentData = diagnostics.latestYear !== null && diagnostics.latestYear >= 2020;
      if (isRecentData) {
        if (
          !bestCandidate ||
          diagnostics.nonZeroMonthCount > bestCandidate.nonZeroMonthCount ||
          (diagnostics.nonZeroMonthCount === bestCandidate.nonZeroMonthCount &&
            diagnostics.recordCount > bestCandidate.recordCount) ||
          (diagnostics.nonZeroMonthCount === bestCandidate.nonZeroMonthCount &&
            diagnostics.recordCount === bestCandidate.recordCount &&
            monthToComparableValue(diagnostics.latestDate) > monthToComparableValue(bestCandidate.latestDate))
        ) {
          bestCandidate = { parsed, sheetName, ...diagnostics };
        }
      }
      return;
    } catch (rowParseError) {
      try {
        let parsed;
        let parserType = "matrix";
        try {
          parsed = buildDataFromExcelDateRows(matrix, sourceFileName, sheetName);
          parserType = "dateRows";
        } catch (dateRowsError) {
          parsed = buildDataFromExcelMatrix(matrix, sourceFileName, sheetName);
          parserType = "matrix";
        }
        const diagnostics = getParsedDataDiagnostics(parsed);
        debugCandidates.push({
          sheetName,
          parser: parserType,
          ...diagnostics,
        });
        const isRecentData = diagnostics.latestYear !== null && diagnostics.latestYear >= 2020;
        if (isRecentData) {
          if (
            !bestCandidate ||
            diagnostics.nonZeroMonthCount > bestCandidate.nonZeroMonthCount ||
            (diagnostics.nonZeroMonthCount === bestCandidate.nonZeroMonthCount &&
              diagnostics.recordCount > bestCandidate.recordCount) ||
            (diagnostics.nonZeroMonthCount === bestCandidate.nonZeroMonthCount &&
              diagnostics.recordCount === bestCandidate.recordCount &&
              monthToComparableValue(diagnostics.latestDate) > monthToComparableValue(bestCandidate.latestDate))
          ) {
            bestCandidate = { parsed, sheetName, ...diagnostics };
          }
        }
        return;
      } catch (matrixParseError) {
        parseErrors.push(`${sheetName}: ${matrixParseError.message}`);
      }
    }
  });

  if (!bestCandidate) {
    const detail = parseErrors.length ? ` (${parseErrors.join(" / ")})` : "";
    throw new Error(`파싱 가능한 시트를 찾지 못했습니다.${detail}`);
  }

  console.group("[Excel Import] Workbook parse result");
  console.log("file:", sourceFileName);
  console.log("sheetNames:", workbook.SheetNames);
  if (debugCandidates.length) {
    console.table(debugCandidates);
  }
  console.log("selectedSheet:", bestCandidate.sheetName);
  console.log("selectedRecordCount:", bestCandidate.recordCount);
  console.groupEnd();

  return bestCandidate.parsed;
}

function bindDataControls() {
  const importExcelButton = document.getElementById("importExcelButton");
  const importExcelInput = document.getElementById("importExcelInput");

  if (!importExcelButton || !importExcelInput) {
    return;
  }

  importExcelButton.addEventListener("click", () => {
    importExcelInput.click();
  });

  importExcelInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      if (!window.XLSX) {
        throw new Error("엑셀 파서 로드에 실패했습니다.");
      }
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, {
        type: "array",
        cellDates: true,
      });
      const parsedData = parseWorkbookToData(workbook, file.name);
      const latestRecord = parsedData.records?.[parsedData.records.length - 1] ?? null;
      console.group("[Excel Import] Selected data snapshot");
      console.log("sourceFile:", parsedData.sourceFile);
      console.log("sheetName:", parsedData.sheetName);
      console.log("recordCount:", parsedData.records?.length ?? 0);
      console.log("latestRecord:", latestRecord);
      console.groupEnd();
      setData(parsedData);
      console.log("[Excel Import] dashboard state", {
        updatedMonth: data.updatedMonth,
        latestNetWorth: latest?.netWorth ?? 0,
        latestInvested: latest?.invested ?? 0,
        latestProfit: latest?.profit ?? 0,
      });
      renderDashboard();
      showFeedback("엑셀 데이터를 불러왔습니다.");
    } catch (error) {
      showFeedback(`엑셀 가져오기 실패: ${error.message}`, true);
    } finally {
      importExcelInput.value = "";
    }
  });
}

function bindTabs() {
  const tabButtons = [...document.querySelectorAll(".tab-button[data-tab-target]")];
  const tabPanels = [...document.querySelectorAll(".tab-panel")];

  if (!tabButtons.length || !tabPanels.length) return;

  function activateTab(targetId) {
    tabButtons.forEach((button) => {
      const isActive = button.dataset.tabTarget === targetId;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    tabPanels.forEach((panel) => {
      panel.classList.toggle("is-active", panel.id === targetId);
    });
  }

  tabButtons.forEach((button, index) => {
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", index === 0 ? "true" : "false");
    button.addEventListener("click", () => {
      activateTab(button.dataset.tabTarget);
    });
  });
}

function getCurrentMonthRecordDate() {
  const now = new Date();
  return buildRecordDate(now.getFullYear(), now.getMonth() + 1);
}

function createInitialEmptyData() {
  const recordDate = getCurrentMonthRecordDate() ?? "2000/01";
  return {
    sourceFile: "초기 상태",
    sheetName: "엑셀 업로드 대기",
    updatedMonth: recordDate,
    records: [
      {
        date: recordDate,
        categories: {},
        investedCategories: {},
        monthlyInvestedCategories: {},
      },
    ],
    latestAllocation: {},
  };
}

setData(createInitialEmptyData());
renderDashboard();
bindDataControls();
bindTabs();

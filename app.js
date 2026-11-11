const BRACKETS = [
  { limitFen: 3600000, rate: 3, quickFen: 0 },
  { limitFen: 14400000, rate: 10, quickFen: 252000 },
  { limitFen: 30000000, rate: 20, quickFen: 1692000 },
  { limitFen: 42000000, rate: 25, quickFen: 3192000 },
  { limitFen: 66000000, rate: 30, quickFen: 5292000 },
  { limitFen: 96000000, rate: 35, quickFen: 8592000 },
  { limitFen: Number.POSITIVE_INFINITY, rate: 45, quickFen: 18192000 }
];

const excelState = {
  workbook: null,
  sheets: [],
  activeEntries: []
};

const monthlyLinkedConfigs = [
  { amountId: "m-pension", rateId: "m-pension-rate", baseId: "m-social-base" },
  { amountId: "m-medical", rateId: "m-medical-rate", baseId: "m-social-base" },
  { amountId: "m-unemployment", rateId: "m-unemployment-rate", baseId: "m-social-base" },
  { amountId: "m-housing", rateId: "m-housing-rate", baseId: "m-housing-base" }
];

const annualLinkedConfigs = [
  { amountId: "a-fill-pension", rateId: "a-fill-pension-rate", baseId: "a-fill-social-base" },
  { amountId: "a-fill-medical", rateId: "a-fill-medical-rate", baseId: "a-fill-social-base" },
  { amountId: "a-fill-unemployment", rateId: "a-fill-unemployment-rate", baseId: "a-fill-social-base" },
  { amountId: "a-fill-housing", rateId: "a-fill-housing-rate", baseId: "a-fill-housing-base" }
];

const annualDetailTopFieldIds = [
  "a-fill-social-base",
  "a-fill-housing-base",
  "a-fill-pension",
  "a-fill-pension-rate",
  "a-fill-medical",
  "a-fill-medical-rate",
  "a-fill-unemployment",
  "a-fill-unemployment-rate",
  "a-fill-housing",
  "a-fill-housing-rate"
];

const annualCommonRowKeys = [
  "targetNet",
  "taxFree",
  "deduction",
  "specialAdd",
  "otherTaxDed",
  "taxReduction"
];

const annualDetailRowKeys = ["pension", "medical", "unemployment", "housing"];
const annualSimpleRowKeys = ["specialDed"];
const annualOptionalRowKeys = ["taxFree", "otherTaxDed", "taxReduction"];

const annualFillSourceIds = {
  targetNet: "a-fill-target-net",
  taxFree: "a-fill-tax-free",
  deduction: "a-fill-deduction",
  specialDed: "a-fill-special-ded",
  specialAdd: "a-fill-special-add",
  pension: "a-fill-pension",
  medical: "a-fill-medical",
  unemployment: "a-fill-unemployment",
  housing: "a-fill-housing",
  otherTaxDed: "a-fill-other-tax-ded",
  taxReduction: "a-fill-tax-reduction"
};

const defaultDeductionText = "5000.00";

const defaultSocialInsurance = {
  socialBase: "4504.00",
  personal: {
    pensionRate: "8",
    medicalRate: "2",
    unemploymentRate: "0.3"
  }
};

const annualFieldOrderConfig = {
  simple: {
    top: [
      "a-mode-net",
      "a-detail-simple",
      "a-fill-target-net",
      "a-fill-deduction",
      "a-fill-special-ded",
      "a-fill-special-add",
      "a-fill-tax-free",
      "a-fill-other-tax-ded",
      "a-fill-tax-reduction"
    ],
    row: [
      "targetNet",
      "deduction",
      "specialDed",
      "specialAdd",
      "taxFree",
      "otherTaxDed",
      "taxReduction"
    ]
  },
  detail: {
    top: [
      "a-mode-net",
      "a-detail-simple",
      "a-fill-target-net",
      "a-fill-deduction",
      "a-fill-special-add",
      "a-fill-social-base",
      "a-fill-housing-base",
      "a-fill-pension",
      "a-fill-pension-rate",
      "a-fill-medical",
      "a-fill-medical-rate",
      "a-fill-unemployment",
      "a-fill-unemployment-rate",
      "a-fill-housing",
      "a-fill-housing-rate",
      "a-fill-tax-free",
      "a-fill-other-tax-ded",
      "a-fill-tax-reduction"
    ],
    row: [
      "targetNet",
      "deduction",
      "specialAdd",
      "pension",
      "medical",
      "unemployment",
      "housing",
      "taxFree",
      "otherTaxDed",
      "taxReduction"
    ]
  }
};

const monthlyState = {
  mode: "gross",
  hiddenAnnuityFen: 0,
  hiddenOtherDirectFen: 0,
  specialDeductionMode: "manual"
};

const annualState = {
  mode: "gross",
  detailMode: "simple"
};

const specialAddHeaders = [
  "累计子女教育支出扣除",
  "累计继续教育支出扣除",
  "累计住房贷款利息支出扣除",
  "累计住房租金支出扣除",
  "累计赡养老人支出扣除",
  "累计3岁以下婴幼儿照护"
];

const otherDeductionHeaders = [
  "累计个人养老金",
  "累计其他扣除",
  "累计准予扣除的捐赠"
];

function parseAmountToFen(rawValue) {
  const value = String(rawValue == null ? "" : rawValue).trim().replace(/,/g, "");
  if (!value) return 0;
  const negative = value.startsWith("-");
  const normalized = negative ? value.slice(1) : value;
  if (!/^\d*(\.\d*)?$/.test(normalized)) {
    throw new Error(`无法识别金额：${rawValue}`);
  }
  const [intPartRaw = "0", fracRaw = ""] = normalized.split(".");
  const intPart = intPartRaw ? Number(intPartRaw) : 0;
  const frac = fracRaw.padEnd(3, "0").slice(0, 3);
  let fen = intPart * 100 + Number(frac.slice(0, 2) || "0");
  if (Number(frac[2]) >= 5) fen += 1;
  return negative ? -fen : fen;
}

function fenToAmountText(fen) {
  const sign = fen < 0 ? "-" : "";
  const abs = Math.abs(fen);
  const intPart = Math.floor(abs / 100);
  const fracPart = String(abs % 100).padStart(2, "0");
  return `${sign}${intPart.toLocaleString("zh-CN")}.${fracPart}`;
}

function fenToInputValue(fen) {
  return fenToAmountText(fen).replace(/,/g, "");
}

function parsePercent(rawValue) {
  const value = String(rawValue == null ? "" : rawValue).trim().replace(/%/g, "");
  if (!value) return null;
  if (!/^\d*(\.\d*)?$/.test(value)) {
    throw new Error(`无法识别比例：${rawValue}`);
  }
  return Number(value);
}

function formatPercent(value) {
  if (!Number.isFinite(value)) return "";
  const fixed = value.toFixed(4);
  return fixed.replace(/\.?0+$/, "");
}

function setText(id, value) {
  document.getElementById(id).textContent = value;
}

function setInputValue(id, fen) {
  document.getElementById(id).value = fenToInputValue(fen);
}

function getBracket(taxableFen) {
  return BRACKETS.find((item) => taxableFen <= item.limitFen);
}

function roundTaxFen(baseFen, rate) {
  return Math.floor((baseFen * rate + 50) / 100);
}

function calculateCumulativeTax(taxableFen) {
  const positiveFen = Math.max(0, taxableFen);
  const bracket = getBracket(positiveFen);
  const grossTaxFen = Math.max(0, roundTaxFen(positiveFen, bracket.rate) - bracket.quickFen);
  return {
    positiveFen,
    rate: bracket.rate,
    quickFen: bracket.quickFen,
    cumulativeTaxFen: grossTaxFen
  };
}

function syncLinkedAmount(config) {
  const baseFen = parseAmountToFen(document.getElementById(config.baseId).value);
  const rate = parsePercent(document.getElementById(config.rateId).value);
  if (!baseFen || rate == null) return;
  setInputValue(config.amountId, Math.round(baseFen * rate / 100));
}

function syncLinkedRate(config) {
  const amountFen = parseAmountToFen(document.getElementById(config.amountId).value);
  const baseFen = parseAmountToFen(document.getElementById(config.baseId).value);
  const rateInput = document.getElementById(config.rateId);
  if (!baseFen) {
    rateInput.value = "";
    return;
  }
  rateInput.value = formatPercent(amountFen * 100 / baseFen);
}

function syncConfigGroup(configs, sourceType, baseId) {
  configs.forEach((config) => {
    if (sourceType === "base" && config.baseId === baseId) {
      syncLinkedAmount(config);
    } else if (sourceType === "rate" && config.rateId === baseId) {
      syncLinkedAmount(config);
    } else if (sourceType === "amount" && config.amountId === baseId) {
      syncLinkedRate(config);
    }
  });
}

function setInputText(id, value) {
  const input = document.getElementById(id);
  if (input) input.value = value;
}

function getDefaultSocialAmountText(rateText) {
  const baseFen = parseAmountToFen(defaultSocialInsurance.socialBase);
  const rate = Number(rateText);
  return fenToInputValue(Math.round(baseFen * rate / 100));
}

function applyLinkedDefault(config, rateText) {
  setInputText(config.rateId, rateText);
  syncLinkedAmount(config);
  document.getElementById(config.amountId).dataset.lastEdited = "rate";
}

function applyMonthlySocialDefaults() {
  setInputText("m-social-base", defaultSocialInsurance.socialBase);
  applyLinkedDefault(monthlyLinkedConfigs[0], defaultSocialInsurance.personal.pensionRate);
  applyLinkedDefault(monthlyLinkedConfigs[1], defaultSocialInsurance.personal.medicalRate);
  applyLinkedDefault(monthlyLinkedConfigs[2], defaultSocialInsurance.personal.unemploymentRate);
}

function applyAnnualSocialDefaults() {
  setInputText("a-fill-social-base", defaultSocialInsurance.socialBase);
  applyLinkedDefault(annualLinkedConfigs[0], defaultSocialInsurance.personal.pensionRate);
  applyLinkedDefault(annualLinkedConfigs[1], defaultSocialInsurance.personal.medicalRate);
  applyLinkedDefault(annualLinkedConfigs[2], defaultSocialInsurance.personal.unemploymentRate);
}

function sumImportedColumns(row, headers) {
  return headers.reduce((sum, header) => sum + parseAmountToFen(row[header]), 0);
}

function extractMonthIndex(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return value.getMonth() + 1;
  }
  const text = String(value == null ? "" : value).trim();
  const matched = text.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (matched) {
    return Math.min(12, Math.max(1, Number(matched[2])));
  }
  const monthOnly = text.match(/(\d{1,2})月/);
  if (monthOnly) {
    return Math.min(12, Math.max(1, Number(monthOnly[1])));
  }
  return 1;
}

function deriveCurrentFromCumulative(totalFen, monthIndex) {
  if (!monthIndex || monthIndex <= 1) return totalFen;
  return Math.round(totalFen / monthIndex);
}

function hasInputText(id) {
  const element = document.getElementById(id);
  return !!String(element && element.value != null ? element.value : "").trim();
}

function resolveCurrentSpecialDeductionFen(model) {
  if (typeof model.currentSpecialDedFen === "number") {
    return model.currentSpecialDedFen;
  }
  return (
    (model.currentPensionFen || 0) +
    (model.currentMedicalFen || 0) +
    (model.currentUnemploymentFen || 0) +
    (model.currentHousingFen || 0) +
    (model.currentAnnuityFen || 0)
  );
}

function getMonthlySpecialDeductionAutoFen() {
  return (
    parseAmountToFen(document.getElementById("m-pension").value) +
    parseAmountToFen(document.getElementById("m-medical").value) +
    parseAmountToFen(document.getElementById("m-unemployment").value) +
    parseAmountToFen(document.getElementById("m-housing").value) +
    monthlyState.hiddenAnnuityFen
  );
}

function hasMonthlySpecialDeductionSourceValue() {
  return (
    monthlyState.hiddenAnnuityFen !== 0 ||
    [
      "m-social-base",
      "m-housing-base",
      "m-pension",
      "m-pension-rate",
      "m-medical",
      "m-medical-rate",
      "m-unemployment",
      "m-unemployment-rate",
      "m-housing",
      "m-housing-rate"
    ].some(hasInputText)
  );
}

function setMonthlySpecialDeductionMode(mode) {
  monthlyState.specialDeductionMode = mode === "auto" ? "auto" : "manual";
}

function buildMonthlyModelFromDom() {
  return {
    month: Number(document.getElementById("m-month").value) || 1,
    inputAmountFen: parseAmountToFen(document.getElementById("m-primary-amount").value),
    mode: monthlyState.mode,
    currentTaxFreeFen: parseAmountToFen(document.getElementById("m-tax-free").value),
    currentDeductionFen: parseAmountToFen(document.getElementById("m-deduction").value),
    currentSpecialAddFen: parseAmountToFen(document.getElementById("m-special-add").value),
    currentSpecialDedFen: parseAmountToFen(document.getElementById("m-current-special-deduction").value),
    currentPensionFen: parseAmountToFen(document.getElementById("m-pension").value),
    currentMedicalFen: parseAmountToFen(document.getElementById("m-medical").value),
    currentUnemploymentFen: parseAmountToFen(document.getElementById("m-unemployment").value),
    currentHousingFen: parseAmountToFen(document.getElementById("m-housing").value),
    currentAnnuityFen: monthlyState.hiddenAnnuityFen,
    currentOtherTaxDedFen: parseAmountToFen(document.getElementById("m-other-tax-ded").value),
    currentOtherDirectFen: monthlyState.hiddenOtherDirectFen,
    currentTaxReductionFen: parseAmountToFen(document.getElementById("m-tax-reduction").value),
    prevIncomeFen: parseAmountToFen(document.getElementById("m-prev-income").value),
    prevTaxFreeFen: parseAmountToFen(document.getElementById("m-prev-tax-free").value),
    prevDeductionFen: parseAmountToFen(document.getElementById("m-prev-deduction").value),
    prevSpecialDedFen: parseAmountToFen(document.getElementById("m-prev-special-ded").value),
    prevSpecialAddFen: parseAmountToFen(document.getElementById("m-prev-special-add").value),
    prevOtherDedFen: parseAmountToFen(document.getElementById("m-prev-other-ded").value),
    prevTaxReductionFen: parseAmountToFen(document.getElementById("m-prev-tax-reduction").value),
    prevTaxPaidFen: parseAmountToFen(document.getElementById("m-prev-tax-paid").value)
  };
}

function calculateMonthly(model, currentGrossFen) {
  const currentSpecialDedFen = resolveCurrentSpecialDeductionFen(model);
  const currentDirectDedFen = currentSpecialDedFen + model.currentOtherDirectFen;
  const cumulativeTaxableFen =
    model.prevIncomeFen +
    currentGrossFen -
    (model.prevTaxFreeFen + model.currentTaxFreeFen) -
    (model.prevDeductionFen + model.currentDeductionFen) -
    (model.prevSpecialDedFen + currentSpecialDedFen) -
    (model.prevSpecialAddFen + model.currentSpecialAddFen) -
    (model.prevOtherDedFen + model.currentOtherTaxDedFen);

  const taxMeta = calculateCumulativeTax(cumulativeTaxableFen);
  const cumulativeTaxAfterReductionFen = Math.max(
    0,
    taxMeta.cumulativeTaxFen - (model.prevTaxReductionFen + model.currentTaxReductionFen)
  );
  const currentTaxFen = Math.max(0, cumulativeTaxAfterReductionFen - model.prevTaxPaidFen);
  const currentNetFen =
    currentGrossFen +
    model.currentTaxFreeFen -
    currentDirectDedFen -
    currentTaxFen;

  return {
    currentGrossFen,
    currentTaxFen,
    currentNetFen,
    cumulativeTaxableFen,
    currentSpecialDedFen,
    currentDirectDedFen,
    rate: taxMeta.rate,
    quickFen: taxMeta.quickFen,
    cumulativeTaxFen: taxMeta.cumulativeTaxFen,
    cumulativeTaxAfterReductionFen
  };
}

function solveMonthlyGross(model) {
  if (model.targetNetFen < 0) {
    throw new Error("目标到手工资不能为负数。");
  }

  const currentSpecialDedFen = resolveCurrentSpecialDeductionFen(model);
  const minimumHighFen =
    model.targetNetFen +
    currentSpecialDedFen +
    model.currentOtherDirectFen +
    500000;

  let low = 0;
  let high = Math.max(1000000, minimumHighFen);

  while (calculateMonthly(model, high).currentNetFen < model.targetNetFen) {
    high *= 2;
    if (high > 1000000000) {
      throw new Error("目标金额过高，超出当前工具的安全测算范围。");
    }
  }

  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (calculateMonthly(model, mid).currentNetFen >= model.targetNetFen) {
      high = mid;
    } else {
      low = mid + 1;
    }
  }

  const result = calculateMonthly(model, low);
  const exact = result.currentNetFen === model.targetNetFen;
  return { result, exact };
}

function updateMonthlySpecialDisplay() {
  const specialInput = document.getElementById("m-current-special-deduction");
  try {
    const totalFen = getMonthlySpecialDeductionAutoFen();
    if (monthlyState.specialDeductionMode === "auto") {
      specialInput.value = hasMonthlySpecialDeductionSourceValue() ? fenToInputValue(totalFen) : "";
    }
  } catch (error) {
    if (monthlyState.specialDeductionMode === "auto") {
      specialInput.value = "";
    }
  }
}

function setMonthlyMode(mode) {
  monthlyState.mode = mode === "gross" ? "gross" : "net";
  document.getElementById("m-mode-gross").classList.toggle("is-active", monthlyState.mode === "gross");
  document.getElementById("m-mode-net").classList.toggle("is-active", monthlyState.mode === "net");
  document.getElementById("m-primary-label").textContent = monthlyState.mode === "gross" ? "税前工资" : "税后工资";
  document.getElementById("m-primary-help").textContent =
    monthlyState.mode === "gross"
      ? "输入本月税前工资后，工具会直接计算本期个税和税后工资。"
      : "输入目标税后工资后，工具会反推本期税前工资。";
  document.getElementById("calc-monthly").textContent = monthlyState.mode === "gross" ? "计算个税" : "反推税前";
}

function calculateTaxBurdenText(incomeFen, taxFen) {
  return `${incomeFen > 0 ? (taxFen * 100 / incomeFen).toFixed(2) : "0.00"}%`;
}

function renderMonthlyResult() {
  const errorEl = document.getElementById("m-error");
  errorEl.textContent = "";

  try {
    updateMonthlySpecialDisplay();
    const model = buildMonthlyModelFromDom();
    if (model.inputAmountFen < 0) {
      throw new Error(model.mode === "gross" ? "税前工资不能为负数。" : "税后工资不能为负数。");
    }
    const calculation = model.mode === "gross"
      ? { result: calculateMonthly(model, model.inputAmountFen), exact: true }
      : solveMonthlyGross({ ...model, targetNetFen: model.inputAmountFen });
    const { result, exact } = calculation;
    const currentYearIncomeFen = model.prevIncomeFen + result.currentGrossFen;
    const currentYearTaxFen = model.prevTaxPaidFen + result.currentTaxFen;
    const taxBurdenText = calculateTaxBurdenText(currentYearIncomeFen, currentYearTaxFen);

    setText("m-result-gross", fenToAmountText(result.currentGrossFen));
    setText("m-result-tax", fenToAmountText(result.currentTaxFen));
    setText("m-result-net", fenToAmountText(result.currentNetFen));
    setText("m-result-taxable", fenToAmountText(Math.max(0, result.cumulativeTaxableFen)));
    setText("m-result-rate", `${result.rate}%`);
    setText("m-result-quick", fenToAmountText(result.quickFen));
    setText("m-result-special", fenToAmountText(result.currentSpecialDedFen));
    setText("m-result-direct", fenToAmountText(result.currentDirectDedFen));
    setText("m-result-tax-burden", taxBurdenText);
    document.getElementById("m-year-tax-burden").value = taxBurdenText;

    const note = model.mode === "gross"
      ? `已按第 ${model.month} 月输入的税前工资直接计算本期个税和税后工资。`
      : exact
        ? `已找到刚好满足目标税后工资的最小税前工资。第 ${model.month} 月填入这个数后，预计到手刚好是 ${fenToAmountText(model.inputAmountFen)}。`
        : "没有找到刚好等于目标税后工资的分值组合，已返回“不低于目标值”的最小税前工资。";
    document.getElementById("m-result-note").textContent = note;

    document.getElementById("m-formula-summary").innerHTML = `
      <strong>${model.mode === "gross" ? "本月工资计算结果" : "本期收入反推结果"}：</strong>${fenToAmountText(result.currentGrossFen)}<br>
      累计应纳税所得额 =
      （上期累计收入 ${fenToAmountText(model.prevIncomeFen)} + 本期收入 ${fenToAmountText(result.currentGrossFen)}）
      - 累计免税收入 ${fenToAmountText(model.prevTaxFreeFen + model.currentTaxFreeFen)}
      - 累计减除费用 ${fenToAmountText(model.prevDeductionFen + model.currentDeductionFen)}
      - 累计专项扣除 ${fenToAmountText(model.prevSpecialDedFen + result.currentSpecialDedFen)}
      - 累计专项附加扣除 ${fenToAmountText(model.prevSpecialAddFen + model.currentSpecialAddFen)}
      - 累计其他扣除 ${fenToAmountText(model.prevOtherDedFen + model.currentOtherTaxDedFen)}<br>
      本期个税 =
      累计税额 ${fenToAmountText(result.cumulativeTaxFen)}
      - 累计减免税额 ${fenToAmountText(model.prevTaxReductionFen + model.currentTaxReductionFen)}
      - 上期已预扣税额 ${fenToAmountText(model.prevTaxPaidFen)}
      ，本期结果为 ${fenToAmountText(result.currentTaxFen)}。<br>
      到手工资 =
      本期收入 ${fenToAmountText(result.currentGrossFen)}
      + 本期免税收入 ${fenToAmountText(model.currentTaxFreeFen)}
      - 本期直接扣款 ${fenToAmountText(result.currentDirectDedFen)}
      - 本期个税 ${fenToAmountText(result.currentTaxFen)}
      = <strong>${fenToAmountText(result.currentNetFen)}</strong><br>
      本年综合税负 =
      （累计已扣税 ${fenToAmountText(currentYearTaxFen)} / 累计收入 ${fenToAmountText(currentYearIncomeFen)}）
      = <strong>${taxBurdenText}</strong>
    `;
  } catch (error) {
    errorEl.textContent = error.message || "计算失败，请检查输入。";
  }
}

function resetMonthlyOutputs() {
  setText("m-result-gross", "0.00");
  setText("m-result-tax", "0.00");
  setText("m-result-net", "0.00");
  setText("m-result-taxable", "0.00");
  setText("m-result-rate", "0%");
  setText("m-result-quick", "0.00");
  setText("m-result-special", "0.00");
  setText("m-result-tax-burden", "0.00%");
  setText("m-result-direct", "0.00");
  document.getElementById("m-year-tax-burden").value = "";
  document.getElementById("m-result-note").textContent = "这里会显示结果说明。";
  document.getElementById("m-formula-summary").textContent = "等待计算。";
  document.getElementById("m-error").textContent = "";
}

function resetMonthly() {
  [
    "m-month",
    "m-primary-amount",
    "m-tax-free",
    "m-special-add",
    "m-current-special-deduction",
    "m-social-base",
    "m-housing-base",
    "m-pension",
    "m-pension-rate",
    "m-medical",
    "m-medical-rate",
    "m-unemployment",
    "m-unemployment-rate",
    "m-housing",
    "m-housing-rate",
    "m-other-tax-ded",
    "m-tax-reduction",
    "m-prev-income",
    "m-prev-tax-free",
    "m-prev-deduction",
    "m-prev-special-ded",
    "m-prev-special-add",
    "m-prev-other-ded",
    "m-prev-tax-reduction",
    "m-prev-tax-paid"
  ].forEach((id) => {
    document.getElementById(id).value = "";
  });
  document.getElementById("m-deduction").value = defaultDeductionText;
  applyMonthlySocialDefaults();
  monthlyState.hiddenAnnuityFen = 0;
  monthlyState.hiddenOtherDirectFen = 0;
  setMonthlySpecialDeductionMode("auto");
  updateMonthlySpecialDisplay();
  setMonthlyMode("gross");
  resetMonthlyOutputs();
}

function buildSheetEntries(sheetName, sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    dateNF: "yyyy-mm-dd"
  });
  const headerIndex = rows.findIndex((row) => {
    const normalized = row.map((cell) => String(cell == null ? "" : cell).trim());
    return normalized.includes("姓名") && normalized.includes("本期收入");
  });
  if (headerIndex < 0) return null;

  const headers = rows[headerIndex].map((cell) => String(cell == null ? "" : cell).trim());
  const entries = [];
  for (let rowIndex = headerIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    if (!row || !row.some((cell) => String(cell == null ? "" : cell).trim())) continue;
    const mapped = {};
    headers.forEach((header, index) => {
      if (header) mapped[header] = row[index];
    });
    const name = String(mapped["姓名"] == null ? "" : mapped["姓名"]).trim();
    const currentIncomeText = String(mapped["本期收入"] == null ? "" : mapped["本期收入"]).trim();
    if (!name && !currentIncomeText) continue;
    const monthIndex = extractMonthIndex(mapped["税款所属期止"] || mapped["税款所属期起"]);
    const periodLabel = String(mapped["税款所属期止"] || mapped["税款所属期起"] || "").trim();
    entries.push({
      id: `${sheetName}-${rowIndex}`,
      rowIndex: rowIndex + 1,
      monthIndex,
      sheetName,
      display: `${name || "未命名"} | ${periodLabel || `${monthIndex}月`} | 本期收入 ${currentIncomeText || "0.00"}`,
      row: mapped
    });
  }

  return {
    name: sheetName,
    headerIndex,
    headers,
    entries
  };
}

function refreshExcelSheetSelect() {
  const sheetSelect = document.getElementById("excel-sheet");
  const rowSelect = document.getElementById("excel-row");
  const fillButton = document.getElementById("excel-fill-monthly");

  if (!excelState.sheets.length) {
    sheetSelect.innerHTML = '<option value="">未识别到可导入工作表</option>';
    sheetSelect.disabled = true;
    rowSelect.innerHTML = '<option value="">请先导入工作表</option>';
    rowSelect.disabled = true;
    fillButton.disabled = true;
    return;
  }

  sheetSelect.disabled = false;
  sheetSelect.innerHTML = excelState.sheets
    .map((sheet, index) => `<option value="${index}">${sheet.name}（${sheet.entries.length} 条）</option>`)
    .join("");
  refreshExcelRowSelect();
}

function refreshExcelRowSelect() {
  const sheetSelect = document.getElementById("excel-sheet");
  const rowSelect = document.getElementById("excel-row");
  const fillButton = document.getElementById("excel-fill-monthly");
  const selectedSheet = excelState.sheets[Number(sheetSelect.value) || 0];

  if (!selectedSheet || !selectedSheet.entries.length) {
    rowSelect.innerHTML = '<option value="">当前工作表没有可导入记录</option>';
    rowSelect.disabled = true;
    fillButton.disabled = true;
    excelState.activeEntries = [];
    return;
  }

  excelState.activeEntries = selectedSheet.entries;
  rowSelect.disabled = false;
  rowSelect.innerHTML = selectedSheet.entries
    .map((entry, index) => `<option value="${index}">${entry.display}</option>`)
    .join("");
  fillButton.disabled = false;
}

async function handleExcelFileChange(event) {
  const file = event.target.files && event.target.files[0];
  const badge = document.getElementById("excel-badge");
  const summary = document.getElementById("excel-summary");

  if (!file) {
    badge.textContent = "未导入文件";
    summary.textContent = "选择 Excel 后，这里会显示识别到的工作表、员工人数和字段匹配情况。";
    return;
  }

  if (typeof XLSX === "undefined") {
    summary.textContent = "Excel 解析库未加载成功，请确认 vendor/xlsx.full.min.js 文件存在。";
    badge.textContent = "解析库未就绪";
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheets = workbook.SheetNames
      .map((sheetName) => buildSheetEntries(sheetName, workbook.Sheets[sheetName]))
      .filter(Boolean);

    excelState.workbook = workbook;
    excelState.sheets = sheets;

    refreshExcelSheetSelect();

    badge.textContent = `${file.name} 已导入`;
    summary.innerHTML = sheets.length
      ? `已识别 <strong>${sheets.length}</strong> 个可导入工作表，共 <strong>${sheets.reduce((sum, sheet) => sum + sheet.entries.length, 0)}</strong> 条员工记录。<br>导入时会自动匹配“姓名、本期收入、累计收入额、累计专项扣除、已缴税额”等字段。对于仅有累计值的专项附加扣除/其他扣除，会按所属月份平均拆分到当月。`
      : "没有识别到包含“姓名 + 本期收入”表头的工作表，请确认文件是工资薪金税款计算表。";
  } catch (error) {
    badge.textContent = "导入失败";
    summary.textContent = error.message || "Excel 导入失败，请确认文件未损坏。";
  }
}

function fillMonthlyFromImportedEntry() {
  const rowSelect = document.getElementById("excel-row");
  const entry = excelState.activeEntries[Number(rowSelect.value) || 0];
  if (!entry) return;

  const row = entry.row;
  const monthIndex = entry.monthIndex || 1;
  const currentIncomeFen = parseAmountToFen(row["本期收入"]);
  const currentTaxFreeFen = parseAmountToFen(row["本期免税收入"]);
  const currentPensionFen = parseAmountToFen(row["本期基本养老保险费"]);
  const currentMedicalFen = parseAmountToFen(row["本期基本医疗保险费"]);
  const currentUnemploymentFen = parseAmountToFen(row["本期失业保险费"]);
  const currentHousingFen = parseAmountToFen(row["本期住房公积金"]);
  const currentAnnuityFen =
    parseAmountToFen(row["本期企业(职业)年金"]) + parseAmountToFen(row["本期企业年金"]);

  const cumulativeIncomeFen = parseAmountToFen(row["累计收入额"]);
  const cumulativeTaxFreeFen = parseAmountToFen(row["累计免税收入"]);
  const cumulativeDeductionFen = parseAmountToFen(row["累计减除费用"]);
  const cumulativeSpecialDedFen = parseAmountToFen(row["累计专项扣除"]);
  const cumulativeSpecialAddFen = sumImportedColumns(row, specialAddHeaders);
  const cumulativeOtherDedFen = sumImportedColumns(row, otherDeductionHeaders);
  const cumulativeTaxReductionFen = parseAmountToFen(row["累计减免税额"]);
  const prevTaxPaidFen = parseAmountToFen(row["已缴税额"]);

  const currentDeductionFen = cumulativeDeductionFen
    ? deriveCurrentFromCumulative(cumulativeDeductionFen, monthIndex)
    : 500000;
  const currentSpecialAddFen = deriveCurrentFromCumulative(cumulativeSpecialAddFen, monthIndex);
  const currentOtherTaxDedFen = deriveCurrentFromCumulative(cumulativeOtherDedFen, monthIndex);
  const currentSpecialDedFen =
    currentPensionFen + currentMedicalFen + currentUnemploymentFen + currentHousingFen + currentAnnuityFen;

  document.getElementById("m-month").value = String(monthIndex);
  document.getElementById("m-tax-free").value = fenToInputValue(currentTaxFreeFen);
  document.getElementById("m-deduction").value = fenToInputValue(currentDeductionFen);
  document.getElementById("m-special-add").value = fenToInputValue(currentSpecialAddFen);
  document.getElementById("m-pension").value = fenToInputValue(currentPensionFen);
  document.getElementById("m-medical").value = fenToInputValue(currentMedicalFen);
  document.getElementById("m-unemployment").value = fenToInputValue(currentUnemploymentFen);
  document.getElementById("m-housing").value = fenToInputValue(currentHousingFen);
  monthlyState.hiddenAnnuityFen = currentAnnuityFen;
  monthlyState.hiddenOtherDirectFen = 0;
  document.getElementById("m-other-tax-ded").value = fenToInputValue(currentOtherTaxDedFen);
  document.getElementById("m-tax-reduction").value = "0.00";
  document.getElementById("m-prev-income").value = fenToInputValue(Math.max(0, cumulativeIncomeFen - currentIncomeFen));
  document.getElementById("m-prev-tax-free").value = fenToInputValue(Math.max(0, cumulativeTaxFreeFen - currentTaxFreeFen));
  document.getElementById("m-prev-deduction").value = fenToInputValue(Math.max(0, cumulativeDeductionFen - currentDeductionFen));
  document.getElementById("m-prev-special-ded").value = fenToInputValue(Math.max(0, cumulativeSpecialDedFen - currentSpecialDedFen));
  document.getElementById("m-prev-special-add").value = fenToInputValue(Math.max(0, cumulativeSpecialAddFen - currentSpecialAddFen));
  document.getElementById("m-prev-other-ded").value = fenToInputValue(Math.max(0, cumulativeOtherDedFen - currentOtherTaxDedFen));
  document.getElementById("m-prev-tax-reduction").value = fenToInputValue(cumulativeTaxReductionFen);
  document.getElementById("m-prev-tax-paid").value = fenToInputValue(prevTaxPaidFen);

  document.getElementById("m-social-base").value = "0.00";
  document.getElementById("m-housing-base").value = "0.00";
  document.getElementById("m-pension-rate").value = "";
  document.getElementById("m-medical-rate").value = "";
  document.getElementById("m-unemployment-rate").value = "";
  document.getElementById("m-housing-rate").value = "";

  const importedModel = {
    month: monthIndex,
    inputAmountFen: currentIncomeFen,
    mode: "gross",
    currentTaxFreeFen,
    currentDeductionFen,
    currentSpecialAddFen,
    currentPensionFen,
    currentMedicalFen,
    currentUnemploymentFen,
    currentHousingFen,
    currentAnnuityFen,
    currentOtherTaxDedFen,
    currentOtherDirectFen: 0,
    currentTaxReductionFen: 0,
    prevIncomeFen: Math.max(0, cumulativeIncomeFen - currentIncomeFen),
    prevTaxFreeFen: Math.max(0, cumulativeTaxFreeFen - currentTaxFreeFen),
    prevDeductionFen: Math.max(0, cumulativeDeductionFen - currentDeductionFen),
    prevSpecialDedFen: Math.max(0, cumulativeSpecialDedFen - currentSpecialDedFen),
    prevSpecialAddFen: Math.max(0, cumulativeSpecialAddFen - currentSpecialAddFen),
    prevOtherDedFen: Math.max(0, cumulativeOtherDedFen - currentOtherTaxDedFen),
    prevTaxReductionFen: cumulativeTaxReductionFen,
    prevTaxPaidFen
  };
  setMonthlyMode("gross");
  document.getElementById("m-primary-amount").value = fenToInputValue(currentIncomeFen);
  setMonthlySpecialDeductionMode("auto");

  document.getElementById("excel-summary").innerHTML =
    `已导入 <strong>${row["姓名"] || "未命名"}</strong> 的 ${monthIndex} 月数据。<br>` +
    `专项附加扣除与累计其他扣除按 ${monthIndex} 个月平均拆分到当月；如果该员工不是按月均匀发生，请在单月参数里手动调整。`;

  updateMonthlySpecialDisplay();
  renderMonthlyResult();
}

function getAnnualFillGrid() {
  return document.getElementById("a-fill-target-net").closest(".field-grid");
}

function getAnnualDetailTopSpecialDeductionFen() {
  return ["a-fill-pension", "a-fill-medical", "a-fill-unemployment", "a-fill-housing"]
    .reduce((sum, id) => sum + parseAmountToFen(document.getElementById(id).value), 0);
}

function hasAnnualDetailTopValue() {
  return annualDetailTopFieldIds.some(hasInputText);
}

function getAnnualRowDetailSpecialDeductionFen(row) {
  return annualDetailRowKeys
    .reduce((sum, key) => sum + parseAmountToFen(row.querySelector(`[data-key="${key}"]`).value), 0);
}

function hasAnnualRowDetailValue(row) {
  return annualDetailRowKeys.some((key) => row.querySelector(`[data-key="${key}"]`).value.trim());
}

function getAnnualRelevantRowKeys() {
  return annualState.detailMode === "detail"
    ? [...annualCommonRowKeys, ...annualDetailRowKeys]
    : [...annualCommonRowKeys, ...annualSimpleRowKeys];
}

function setVisualOrder(element, order) {
  if (!element) return;
  element.style.order = order == null ? "" : String(order);
}

function updateAnnualFieldOrder() {
  const orderConfig = annualFieldOrderConfig[annualState.detailMode] || annualFieldOrderConfig.simple;
  orderConfig.top.forEach((id, index) => {
    setVisualOrder(document.getElementById(id)?.closest(".field"), index + 1);
  });

  document.querySelectorAll("#annual-body .annual-item").forEach((row) => {
    orderConfig.row.forEach((key, index) => {
      setVisualOrder(row.querySelector(`[data-key="${key}"]`)?.closest(".field-inline"), index + 1);
    });
  });
}

function syncAnnualSimpleSpecialDeductionFromDetail() {
  const topInput = document.getElementById("a-fill-special-ded");
  if (topInput && !topInput.value.trim() && hasAnnualDetailTopValue()) {
    try {
      topInput.value = fenToInputValue(getAnnualDetailTopSpecialDeductionFen());
    } catch (error) {
      // Ignore transient parse errors while editing.
    }
  }

  document.querySelectorAll("#annual-body .annual-item").forEach((row) => {
    const specialInput = row.querySelector('[data-key="specialDed"]');
    if (!specialInput || specialInput.value.trim() || !hasAnnualRowDetailValue(row)) return;
    try {
      specialInput.value = fenToInputValue(getAnnualRowDetailSpecialDeductionFen(row));
    } catch (error) {
      // Ignore transient parse errors while editing.
    }
  });
}

function isZeroOrEmptyAmountText(text) {
  const value = String(text == null ? "" : text).trim();
  if (!value) return true;
  try {
    return parseAmountToFen(value) === 0;
  } catch (error) {
    return false;
  }
}

function shouldHideAnnualOptionalField(key) {
  const fillElement = document.getElementById(annualFillSourceIds[key]);
  return isZeroOrEmptyAmountText(fillElement && fillElement.value);
}

function updateAnnualOptionalRowFieldVisibility(rows = Array.from(document.querySelectorAll("#annual-body .annual-item"))) {
  annualOptionalRowKeys.forEach((key) => {
    const hidden = shouldHideAnnualOptionalField(key);
    rows.forEach((row) => {
      const field = row.querySelector(`[data-key="${key}"]`)?.closest(".field-inline");
      if (field) {
        field.hidden = hidden;
      }
    });
  });
}

function resetAnnualOptionalRowFieldVisibility(rows = Array.from(document.querySelectorAll("#annual-body .annual-item"))) {
  rows.forEach((row) => {
    annualOptionalRowKeys.forEach((key) => {
      const field = row.querySelector(`[data-key="${key}"]`)?.closest(".field-inline");
      if (field) {
        field.hidden = false;
      }
    });
  });
}



function ensureAnnualModeControls() {
  const targetInput = document.getElementById("a-fill-target-net");
  const targetField = targetInput.closest(".field");
  const targetLabel = targetField.querySelector("label");
  targetLabel.id = "a-fill-primary-label";
  if (document.getElementById("a-mode-net")) return;

  const modeField = document.createElement("div");
  modeField.className = "field";
  modeField.innerHTML =
    '<label>\u7edf\u4e00\u8f93\u5165\u53e3\u5f84</label>' +
    '<div class="mode-switch" role="tablist" aria-label="\u5168\u5e74\u6279\u91cf\u8bbe\u7f6e\u8f93\u5165\u53e3\u5f84">' +
      '<button type="button" class="mode-button is-active" id="a-mode-gross" data-annual-mode="gross">\u7edf\u4e00\u7a0e\u524d\u6536\u5165</button>' +
      '<button type="button" class="mode-button" id="a-mode-net" data-annual-mode="net">\u7edf\u4e00\u5230\u624b\u76ee\u6807</button>' +
    '</div>';
  targetField.parentElement.insertBefore(modeField, targetField);
}


function ensureAnnualDetailControls() {
  const annualPanel = document.getElementById("annual-panel");
  const fillGrid = getAnnualFillGrid();
  annualPanel.dataset.annualDetailMode = annualState.detailMode;

  annualDetailTopFieldIds.forEach((id) => {
    const field = document.getElementById(id).closest(".field");
    field.dataset.annualSection = "detail-special";
  });

  if (!document.getElementById("a-fill-special-ded")) {
    const deductionField = document.getElementById("a-fill-deduction").closest(".field");
    const specialField = document.createElement("div");
    specialField.className = "field";
    specialField.dataset.annualSection = "simple-special";
    specialField.innerHTML =
      '<label for="a-fill-special-ded">\u7edf\u4e00\u4e13\u9879\u6263\u9664</label>' +
      '<input id="a-fill-special-ded" type="text" inputmode="decimal" value="">';
    deductionField.insertAdjacentElement("afterend", specialField);
  }

  if (!document.getElementById("a-detail-simple")) {
    const layoutField = document.createElement("div");
    layoutField.className = "field";
    layoutField.dataset.annualSection = "layout-toggle";
    layoutField.innerHTML =
      '<label>\u4e13\u9879\u6263\u9664\u586b\u5199\u65b9\u5f0f</label>' +
      '<div class="mode-switch" role="tablist" aria-label="\u5168\u5e74\u4e13\u9879\u6263\u9664\u586b\u5199\u65b9\u5f0f">' +
        '<button type="button" class="mode-button is-active" id="a-detail-simple" data-annual-detail-mode="simple">\u7cbe\u7b80\u6a21\u5f0f</button>' +
        '<button type="button" class="mode-button" id="a-detail-detail" data-annual-detail-mode="detail">\u8be6\u7ec6\u6a21\u5f0f</button>' +
      '</div>';
    const annualModeField = document.getElementById("a-mode-net").closest(".field");
    annualModeField.insertAdjacentElement("afterend", layoutField);
  }

  const annualCalcTitle = annualPanel.querySelectorAll(".card h2")[1];
  if (annualCalcTitle) {
    const annualCalcTitleText = annualCalcTitle.querySelector("span");
    if (annualCalcTitleText) {
      annualCalcTitleText.textContent = "\u9010\u6708\u8ba1\u7b97";
    } else {
      annualCalcTitle.textContent = "\u9010\u6708\u8ba1\u7b97";
    }
  }

  fillGrid.dataset.annualFillGrid = "ready";
}

function setAnnualPrimaryResultLabel(row, text) {
  const primaryResult = row.querySelector(".annual-result");
  const strong = primaryResult.querySelector("strong");
  if (primaryResult.firstChild && primaryResult.firstChild.nodeType === Node.TEXT_NODE) {
    primaryResult.firstChild.textContent = text;
  } else {
    primaryResult.insertBefore(document.createTextNode(text), strong);
  }
}



function updateAnnualModeLabels() {
  const isGross = annualState.mode === "gross";
  const primaryLabel = document.getElementById("a-fill-primary-label");
  if (primaryLabel) {
    primaryLabel.textContent = isGross ? "\u7edf\u4e00\u7a0e\u524d\u6536\u5165" : "\u7edf\u4e00\u5230\u624b\u76ee\u6807";
  }
  const netButton = document.getElementById("a-mode-net");
  const grossButton = document.getElementById("a-mode-gross");
  if (netButton && grossButton) {
    netButton.classList.toggle("is-active", !isGross);
    grossButton.classList.toggle("is-active", isGross);
  }
  document.querySelectorAll("#annual-body .annual-item").forEach((row) => {
    const primaryInputLabel = row.querySelector(".annual-item-grid .field-inline label");
    if (primaryInputLabel) {
      primaryInputLabel.textContent = isGross ? "\u7a0e\u524d\u6536\u5165" : "\u76ee\u6807\u5230\u624b";
    }
    setAnnualPrimaryResultLabel(row, isGross ? "\u672c\u6708\u5230\u624b" : "\u53cd\u63a8\u7a0e\u524d");
  });
}

function updateAnnualDetailModeUI() {
  const annualPanel = document.getElementById("annual-panel");
  annualPanel.dataset.annualDetailMode = annualState.detailMode;

  const simpleButton = document.getElementById("a-detail-simple");
  const detailButton = document.getElementById("a-detail-detail");
  if (simpleButton && detailButton) {
    simpleButton.classList.toggle("is-active", annualState.detailMode === "simple");
    detailButton.classList.toggle("is-active", annualState.detailMode === "detail");
  }
}

function setAnnualMode(mode) {
  annualState.mode = mode === "gross" ? "gross" : "net";
  updateAnnualModeLabels();
}

function setAnnualDetailMode(mode) {
  annualState.detailMode = mode === "detail" ? "detail" : "simple";
  if (annualState.detailMode === "simple") {
    syncAnnualSimpleSpecialDeductionFromDetail();
  }
  updateAnnualDetailModeUI();
  updateAnnualFieldOrder();
}



function resetAnnualOutputs() {
  document.querySelectorAll("#annual-body .annual-item").forEach((row) => {
    row.querySelector('[data-out="gross"]').textContent = "0.00";
    row.querySelector('[data-out="tax"]').textContent = "0.00";
    row.querySelector('[data-out="net"]').textContent = "0.00";
    row.querySelector('[data-out="taxPaid"]').textContent = "0.00";
    row.querySelector('[data-out="taxPaidText"]').textContent = "0.00";
  });
  setText("a-sum-gross", "0.00");
  setText("a-sum-tax", "0.00");
  setText("a-sum-net", "0.00");
  setText("a-sum-special", "0.00");
  setText("a-sum-special-add", "0.00");
  setText("a-sum-tax-burden", "0.00%");
  document.getElementById("a-summary").textContent = "\u7b49\u5f85\u8ba1\u7b97\u3002";
  document.getElementById("a-error").textContent = "";
}

function setupLinkedInputs(configs, onChange) {
  const baseIds = [...new Set(configs.map((config) => config.baseId))];

  configs.forEach((config) => {
    document.getElementById(config.amountId).addEventListener("input", () => {
      document.getElementById(config.amountId).dataset.lastEdited = "amount";
      try {
        syncLinkedRate(config);
      } catch (error) {
        // Keep manual input while user is typing.
      }
      if (onChange) onChange();
    });

    document.getElementById(config.rateId).addEventListener("input", () => {
      document.getElementById(config.amountId).dataset.lastEdited = "rate";
      try {
        syncLinkedAmount(config);
      } catch (error) {
        // Ignore transient parse errors while editing.
      }
      if (onChange) onChange();
    });
  });

  baseIds.forEach((baseId) => {
    document.getElementById(baseId).addEventListener("input", () => {
      configs
        .filter((config) => config.baseId === baseId)
        .forEach((config) => {
          try {
            const amountInput = document.getElementById(config.amountId);
            const rateInput = document.getElementById(config.rateId);
            if ((amountInput.dataset.lastEdited || "") === "amount" || !rateInput.value.trim()) {
              syncLinkedRate(config);
            } else {
              syncLinkedAmount(config);
            }
          } catch (error) {
            // Ignore transient parse errors while editing.
          }
        });
      if (onChange) onChange();
    });
  });
}

function setupMonthlySpecialDeductionEvents() {
  const specialField = document.getElementById("m-current-special-deduction");
  const specialHint = specialField.closest(".field").querySelector("small");
  specialHint.textContent = "可手动填列，也可根据社保公积金相关数值自动计算。";
  specialField.addEventListener("input", () => {
    setMonthlySpecialDeductionMode("manual");
  });

  const sourceIds = new Set([
    "m-social-base",
    "m-housing-base",
    ...monthlyLinkedConfigs.flatMap((config) => [config.amountId, config.rateId, config.baseId])
  ]);
  sourceIds.forEach((id) => {
    document.getElementById(id).addEventListener("input", () => {
      setMonthlySpecialDeductionMode("auto");
      updateMonthlySpecialDisplay();
    });
  });
}

function setupMonthlyLayoutOrder() {
  document.getElementById("m-error").closest(".card").classList.add("monthly-result-card");
}

function setupTabs() {
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab-button").forEach((item) => item.classList.remove("is-active"));
      document.querySelectorAll(".panel").forEach((panel) => panel.classList.remove("is-active"));
      button.classList.add("is-active");
      document.getElementById(button.dataset.tabTarget).classList.add("is-active");
    });
  });
}

function createAnnualRows() {
  const tbody = document.getElementById("annual-body");
  const rows = [];
  for (let month = 1; month <= 12; month += 1) {
    rows.push(`
      <div class="annual-item" data-month="${month}">
        <div class="annual-item-head">
          <strong>${month} 月</strong>
          <span class="import-badge" data-out="taxPaid">0.00</span>
        </div>
        <div class="annual-item-grid">
          <div class="field-inline"><label>目标到手</label><input data-key="targetNet" type="text" inputmode="decimal" value="20000.00"></div>
          <div class="field-inline"><label>免税收入</label><input data-key="taxFree" type="text" inputmode="decimal" value="0.00"></div>
          <div class="field-inline"><label>减除费用</label><input data-key="deduction" type="text" inputmode="decimal" value="5000.00"></div>
          <div class="field-inline" data-annual-section="simple-special"><label>专项扣除</label><input data-key="specialDed" type="text" inputmode="decimal" value=""></div>
          <div class="field-inline"><label>专项附加</label><input data-key="specialAdd" type="text" inputmode="decimal" value="0.00"></div>
          <div class="field-inline" data-annual-section="detail-special"><label>养老</label><input data-key="pension" type="text" inputmode="decimal" value="${getDefaultSocialAmountText(defaultSocialInsurance.personal.pensionRate)}"></div>
          <div class="field-inline" data-annual-section="detail-special"><label>医疗</label><input data-key="medical" type="text" inputmode="decimal" value="${getDefaultSocialAmountText(defaultSocialInsurance.personal.medicalRate)}"></div>
          <div class="field-inline" data-annual-section="detail-special"><label>失业</label><input data-key="unemployment" type="text" inputmode="decimal" value="${getDefaultSocialAmountText(defaultSocialInsurance.personal.unemploymentRate)}"></div>
          <div class="field-inline" data-annual-section="detail-special"><label>公积金</label><input data-key="housing" type="text" inputmode="decimal" value="0.00"></div>
          <div class="field-inline"><label>其他税前扣除</label><input data-key="otherTaxDed" type="text" inputmode="decimal" value="0.00"></div>
          <div class="field-inline"><label>减免税额</label><input data-key="taxReduction" type="text" inputmode="decimal" value="0.00"></div>
        </div>
        <div class="annual-result-strip">
          <div class="annual-result">反推税前<strong data-out="gross">0.00</strong></div>
          <div class="annual-result">本月个税<strong data-out="tax">0.00</strong></div>
          <div class="annual-result">税后工资<strong data-out="net">0.00</strong></div>
          <div class="annual-result">累计已扣税<strong data-out="taxPaidText">0.00</strong></div>
        </div>
      </div>
    `);
  }
  tbody.innerHTML = rows.join("");
}

function getAnnualRowInputs(row) {
  const read = (key) => parseAmountToFen(row.querySelector(`[data-key="${key}"]`).value);
  return {
    primaryAmountFen: read("targetNet"),
    currentTaxFreeFen: read("taxFree"),
    currentDeductionFen: read("deduction"),
    currentSpecialDedFen: read("specialDed"),
    currentPensionFen: read("pension"),
    currentMedicalFen: read("medical"),
    currentUnemploymentFen: read("unemployment"),
    currentHousingFen: read("housing"),
    currentSpecialAddFen: read("specialAdd"),
    currentOtherTaxDedFen: read("otherTaxDed"),
    currentTaxReductionFen: read("taxReduction")
  };
}



function renderAnnual() {
  const errorEl = document.getElementById("a-error");
  errorEl.textContent = "";

  try {
    const relevantKeys = getAnnualRelevantRowKeys();
    let prevIncomeFen = 0;
    let prevTaxFreeFen = 0;
    let prevDeductionFen = 0;
    let prevSpecialDedFen = 0;
    let prevSpecialAddFen = 0;
    let prevOtherDedFen = 0;
    let prevTaxReductionFen = 0;
    let prevTaxPaidFen = 0;

    let sumGrossFen = 0;
    let sumTaxFen = 0;
    let sumNetFen = 0;
    let sumSpecialFen = 0;
    let sumSpecialAddFen = 0;

    const rows = Array.from(document.querySelectorAll("#annual-body .annual-item"));
    updateAnnualOptionalRowFieldVisibility(rows);

    rows.forEach((row) => {
      const current = getAnnualRowInputs(row);
      const hasAnyInput = relevantKeys.some((key) => row.querySelector(`[data-key="${key}"]`).value.trim());
      if (!hasAnyInput) {
        row.querySelector('[data-out="gross"]').textContent = "0.00";
        row.querySelector('[data-out="tax"]').textContent = "0.00";
        row.querySelector('[data-out="net"]').textContent = "0.00";
        row.querySelector('[data-out="taxPaid"]').textContent = "0.00";
        row.querySelector('[data-out="taxPaidText"]').textContent = "0.00";
        return;
      }

      const isGrossMode = annualState.mode === "gross";
      const isSimpleDetailMode = annualState.detailMode === "simple";
      if (current.primaryAmountFen < 0) {
        throw new Error(isGrossMode ? "\u7a0e\u524d\u6536\u5165\u4e0d\u80fd\u4e3a\u8d1f\u6570\u3002" : "\u5230\u624b\u76ee\u6807\u4e0d\u80fd\u4e3a\u8d1f\u6570\u3002");
      }

      const model = {
        month: Number(row.dataset.month),
        currentTaxFreeFen: current.currentTaxFreeFen,
        currentDeductionFen: current.currentDeductionFen,
        currentSpecialAddFen: current.currentSpecialAddFen,
        currentPensionFen: isSimpleDetailMode ? 0 : current.currentPensionFen,
        currentMedicalFen: isSimpleDetailMode ? 0 : current.currentMedicalFen,
        currentUnemploymentFen: isSimpleDetailMode ? 0 : current.currentUnemploymentFen,
        currentHousingFen: isSimpleDetailMode ? 0 : current.currentHousingFen,
        currentAnnuityFen: 0,
        currentOtherTaxDedFen: current.currentOtherTaxDedFen,
        currentOtherDirectFen: 0,
        currentTaxReductionFen: current.currentTaxReductionFen,
        prevIncomeFen,
        prevTaxFreeFen,
        prevDeductionFen,
        prevSpecialDedFen,
        prevSpecialAddFen,
        prevOtherDedFen,
        prevTaxReductionFen,
        prevTaxPaidFen
      };
      if (isSimpleDetailMode) {
        model.currentSpecialDedFen = current.currentSpecialDedFen;
      }

      const result = isGrossMode
        ? calculateMonthly(model, current.primaryAmountFen)
        : solveMonthlyGross({ ...model, targetNetFen: current.primaryAmountFen }).result;

      row.querySelector('[data-out="gross"]').textContent = fenToAmountText(
        isGrossMode ? result.currentNetFen : result.currentGrossFen
      );
      row.querySelector('[data-out="tax"]').textContent = fenToAmountText(result.currentTaxFen);
      row.querySelector('[data-out="net"]').textContent = fenToAmountText(result.currentNetFen);

      prevIncomeFen += result.currentGrossFen;
      prevTaxFreeFen += current.currentTaxFreeFen;
      prevDeductionFen += current.currentDeductionFen;
      prevSpecialDedFen += result.currentSpecialDedFen;
      prevSpecialAddFen += current.currentSpecialAddFen;
      prevOtherDedFen += current.currentOtherTaxDedFen;
      prevTaxReductionFen += current.currentTaxReductionFen;
      prevTaxPaidFen += result.currentTaxFen;

      row.querySelector('[data-out="taxPaid"]').textContent = `${result.rate}% / ${fenToAmountText(result.quickFen)}`;
      row.querySelector('[data-out="taxPaidText"]').textContent = fenToAmountText(prevTaxPaidFen);

      sumGrossFen += result.currentGrossFen;
      sumTaxFen += result.currentTaxFen;
      sumNetFen += result.currentNetFen;
      sumSpecialFen += result.currentSpecialDedFen;
      sumSpecialAddFen += current.currentSpecialAddFen;
    });

    setText("a-sum-gross", fenToAmountText(sumGrossFen));
    setText("a-sum-tax", fenToAmountText(sumTaxFen));
    setText("a-sum-net", fenToAmountText(sumNetFen));
    setText("a-sum-special", fenToAmountText(sumSpecialFen));
    setText("a-sum-special-add", fenToAmountText(sumSpecialAddFen));
    setText(
      "a-sum-tax-burden",
      `${sumGrossFen > 0 ? (sumTaxFen * 100 / sumGrossFen).toFixed(2) : "0.00"}%`
    );

    document.getElementById("a-summary").innerHTML = `
      <strong>\u5e74\u5ea6\u6a21\u5f0f\u8bf4\u660e\uff1a</strong>\u8fd9\u91cc\u6309 1 \u6708\u5230 12 \u6708\u4f9d\u6b21\u7d2f\u8ba1\u9884\u6263\u9884\u7f34\uff0c\u4e0a\u4e00\u6708\u7ed3\u679c\u4f1a\u81ea\u52a8\u6eda\u5165\u4e0b\u4e00\u6708\u3002
      \u5168\u5e74\u7a0e\u524d\u6536\u5165\u4e3a ${fenToAmountText(sumGrossFen)}\uff0c\u5168\u5e74\u4e2a\u7a0e\u4e3a ${fenToAmountText(sumTaxFen)}\uff0c\u5168\u5e74\u9884\u8ba1\u5230\u624b\u4e3a <strong>${fenToAmountText(sumNetFen)}</strong>\u3002
      \u8fd9\u91cc\u7684\u201c\u5168\u5e74\u7efc\u5408\u7a0e\u8d1f\u201d\u6309 <strong>\u5168\u5e74\u4e2a\u7a0e / \u5168\u5e74\u7a0e\u524d\u6536\u5165</strong> \u8ba1\u7b97\u3002
    `;
  } catch (error) {
    errorEl.textContent = error.message || "\u5e74\u5ea6\u8ba1\u7b97\u5931\u8d25\uff0c\u8bf7\u68c0\u67e5\u8f93\u5165\u3002";
  }
}

function getAnnualActiveFillKeys() {
  return annualState.detailMode === "detail"
    ? [...annualCommonRowKeys, ...annualDetailRowKeys]
    : [...annualCommonRowKeys, ...annualSimpleRowKeys];
}

function applyAnnualFill() {
  if (annualState.detailMode === "simple") {
    syncAnnualSimpleSpecialDeductionFromDetail();
  }

  const fillKeys = getAnnualActiveFillKeys();
  const fillValues = Object.fromEntries(
    fillKeys.map((key) => [key, document.getElementById(annualFillSourceIds[key]).value])
  );

  document.querySelectorAll("#annual-body .annual-item").forEach((row) => {
    fillKeys.forEach((key) => {
      row.querySelector(`[data-key="${key}"]`).value = fillValues[key];
    });
  });
  renderAnnual();
}

function resetAnnual() {
  setAnnualMode("gross");
  document.getElementById("a-fill-target-net").value = "";
  document.getElementById("a-fill-tax-free").value = "";
  document.getElementById("a-fill-special-ded").value = "";
  document.getElementById("a-fill-social-base").value = "";
  document.getElementById("a-fill-housing-base").value = "";
  document.getElementById("a-fill-pension").value = "";
  document.getElementById("a-fill-pension-rate").value = "";
  document.getElementById("a-fill-medical").value = "";
  document.getElementById("a-fill-medical-rate").value = "";
  document.getElementById("a-fill-unemployment").value = "";
  document.getElementById("a-fill-unemployment-rate").value = "";
  document.getElementById("a-fill-housing").value = "";
  document.getElementById("a-fill-housing-rate").value = "";
  document.getElementById("a-fill-special-add").value = "";
  document.getElementById("a-fill-other-tax-ded").value = "";
  document.getElementById("a-fill-tax-reduction").value = "";
  document.getElementById("a-fill-deduction").value = defaultDeductionText;
  applyAnnualSocialDefaults();
  document.querySelectorAll("#annual-body .annual-item").forEach((row) => {
    row.querySelectorAll("input").forEach((input) => {
      input.value = "";
    });
    row.querySelector('[data-key="deduction"]').value = defaultDeductionText;
    row.querySelector('[data-key="pension"]').value = getDefaultSocialAmountText(defaultSocialInsurance.personal.pensionRate);
    row.querySelector('[data-key="medical"]').value = getDefaultSocialAmountText(defaultSocialInsurance.personal.medicalRate);
    row.querySelector('[data-key="unemployment"]').value = getDefaultSocialAmountText(defaultSocialInsurance.personal.unemploymentRate);
  });
  resetAnnualOptionalRowFieldVisibility();
  resetAnnualOutputs();
}

function triggerPressFeedback(element) {
  if (!element || element.disabled) return;
  element.classList.remove("is-pressed");
  if (element._pressFeedbackTimer) {
    window.clearTimeout(element._pressFeedbackTimer);
  }
  void element.offsetWidth;
  element.classList.add("is-pressed");
  element._pressFeedbackTimer = window.setTimeout(() => {
    element.classList.remove("is-pressed");
    element._pressFeedbackTimer = null;
  }, 180);
}

function setupInteractiveFeedback() {
  document.addEventListener("click", (event) => {
    const interactive = event.target.closest('button.primary, button.secondary, button.ghost');
    if (!interactive) return;
    triggerPressFeedback(interactive);
  });
}

function updateCollapseToggleLabel(button, isCollapsed) {
  const sectionTitle = button.dataset.sectionTitle || "区域";
  button.setAttribute("aria-expanded", String(!isCollapsed));
  button.setAttribute("aria-label", `${isCollapsed ? "展开" : "收起"}${sectionTitle}`);
}

function setupCollapsibleCards() {
  document.querySelectorAll("[data-collapse-toggle]").forEach((button) => {
    const card = button.closest(".collapsible-card");
    if (!card) return;

    updateCollapseToggleLabel(button, card.classList.contains("is-collapsed"));
  });

  if (setupCollapsibleCards.bound) return;
  document.addEventListener("click", (event) => {
    const button = event.target.closest("[data-collapse-toggle]");
    const heading = button ? null : event.target.closest(".collapsible-card > h2");
    const toggle = button || (heading && heading.querySelector("[data-collapse-toggle]"));
    if (!toggle) return;

    const card = toggle.closest(".collapsible-card");
    if (!card) return;

    const isCollapsed = card.classList.toggle("is-collapsed");
    updateCollapseToggleLabel(toggle, isCollapsed);
  });
  setupCollapsibleCards.bound = true;
}

function setupEvents() {
  document.getElementById("calc-monthly").addEventListener("click", renderMonthlyResult);
  document.getElementById("reset-monthly").addEventListener("click", resetMonthly);
  document.querySelectorAll("[data-monthly-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      setMonthlyMode(button.dataset.monthlyMode);
      renderMonthlyResult();
    });
  });
  document.getElementById("excel-file").addEventListener("change", handleExcelFileChange);
  document.getElementById("excel-sheet").addEventListener("change", refreshExcelRowSelect);
  document.getElementById("excel-fill-monthly").addEventListener("click", fillMonthlyFromImportedEntry);
  setupMonthlySpecialDeductionEvents();

  setupLinkedInputs(monthlyLinkedConfigs, updateMonthlySpecialDisplay);
  setupLinkedInputs(annualLinkedConfigs);
  document.querySelectorAll("[data-annual-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      setAnnualMode(button.dataset.annualMode);
      renderAnnual();
    });
  });
  document.querySelectorAll("[data-annual-detail-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      setAnnualDetailMode(button.dataset.annualDetailMode);
      renderAnnual();
    });
  });

  document.getElementById("apply-annual-fill").addEventListener("click", applyAnnualFill);
  document.getElementById("calc-annual").addEventListener("click", renderAnnual);
  document.getElementById("reset-annual").addEventListener("click", resetAnnual);
  document.getElementById("annual-body").addEventListener("input", renderAnnual);
  setupInteractiveFeedback();
  setupCollapsibleCards();
}

ensureAnnualModeControls();
ensureAnnualDetailControls();
createAnnualRows();
setupMonthlyLayoutOrder();
setupTabs();
setupEvents();
setMonthlyMode("gross");
setAnnualMode("gross");
setAnnualDetailMode("simple");
document.getElementById("m-deduction").value = defaultDeductionText;
document.getElementById("a-fill-deduction").value = defaultDeductionText;
resetMonthly();
resetAnnual();

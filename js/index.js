/* =========================
   Homepage interactions + data monitor
========================= */

(() => {
    "use strict";

    const config = window.APP_CONFIG || {};
    const ICHI_EXCEL_URL = config.ICHI_EXCEL_URL || "ichi data.xlsx";
    const FISCAL_EXCEL_URL = config.FISCAL_EXCEL_URL || "fiscal data.xlsx";

    /* ---------- Mobile menu ---------- */

    const menuButton = document.querySelector(".menu-button");
    const mobileNav = document.querySelector(".mobile-nav");

    if (menuButton && mobileNav) {
        menuButton.addEventListener("click", () => {
            mobileNav.classList.toggle("open");

            menuButton.setAttribute(
                "aria-expanded",
                mobileNav.classList.contains("open") ? "true" : "false"
            );
        });
    }

    /* ---------- Data tabs ---------- */

    const dataButtons = document.querySelectorAll(".data-switch-btn");
    const dataPanels = document.querySelectorAll(".data-panel");

    dataButtons.forEach(button => {
        button.addEventListener("click", () => {
            const target = button.dataset.target;

            dataButtons.forEach(btn => btn.classList.remove("active"));
            dataPanels.forEach(panel => panel.classList.remove("active"));

            button.classList.add("active");

            const targetPanel = document.getElementById(target);
            if (targetPanel) {
                targetPanel.classList.add("active");
            }
        });
    });

    /* ---------- Generic Excel helpers ---------- */

    function numberValue(value) {
        if (value === null || value === undefined || value === "") {
            return null;
        }

        const n = Number(value);
        return Number.isFinite(n) ? n : null;
    }

    function cacheBust(url) {
        const separator = url.includes("?") ? "&" : "?";
        return `${url}${separator}v=${Date.now()}`;
    }

    async function fetchWorkbook(url) {
        if (typeof XLSX === "undefined") {
            throw new Error("SheetJS 未加载");
        }

        const response = await fetch(cacheBust(encodeURI(url)), {
            cache: "no-store"
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const buffer = await response.arrayBuffer();

        return XLSX.read(buffer, {
            type: "array",
            cellDates: false
        });
    }

    function findHeader(rows, requiredHeaders) {
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
            const row = (rows[i] || []).map(value =>
                String(value ?? "").trim()
            );

            if (requiredHeaders.every(header => row.includes(header))) {
                return i;
            }
        }

        return -1;
    }

    function excelMonth(value) {
        if (value === null || value === undefined || value === "") {
            return null;
        }

        if (typeof value === "number" && XLSX?.SSF?.parse_date_code) {
            const parts = XLSX.SSF.parse_date_code(value);

            if (parts?.y && parts?.m) {
                return `${parts.y}-${String(parts.m).padStart(2, "0")}`;
            }
        }

        const text = String(value).trim();
        const match = text.match(/^(\d{4})[-/.](\d{1,2})/);

        if (match) {
            return `${match[1]}-${String(match[2]).padStart(2, "0")}`;
        }

        const date = new Date(text);

        if (!Number.isNaN(date.getTime())) {
            return `${date.getFullYear()}-${String(
                date.getMonth() + 1
            ).padStart(2, "0")}`;
        }

        return null;
    }

    function setMetric(id, value, suffix = "") {
        const element = document.getElementById(id);
        if (!element) return;

        if (value === null || value === undefined || !Number.isFinite(Number(value))) {
            element.textContent = "--";
            element.classList.remove("negative");
            return;
        }

        const number = Number(value);
        element.textContent = `${number.toFixed(2)}${suffix}`;
        element.classList.toggle("negative", number < 0);
    }

    function setText(id, text) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = text || "--";
        }
    }

    /* ---------- ICHI homepage snapshot ---------- */

    async function loadICHIHomepage() {
        const workbook = await fetchWorkbook(ICHI_EXCEL_URL);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: true,
            defval: null
        });

        const required = [
            "weeknum",
            "消费景气指数",
            "投资景气指数",
            "生产景气指数",
            "综合景气指数"
        ];

        const headerRow = findHeader(rows, required);

        if (headerRow < 0) {
            throw new Error("ICHI Excel 未找到所需表头");
        }

        const headers = (rows[headerRow] || []).map(value =>
            String(value ?? "").trim()
        );

        const indexOf = name => headers.indexOf(name);

        let latest = null;

        for (let i = headerRow + 1; i < rows.length; i++) {
            const row = rows[i] || [];
            const weeknum = String(row[indexOf("weeknum")] ?? "").trim();

            if (!weeknum) continue;

            const item = {
                weeknum,
                composite: numberValue(row[indexOf("综合景气指数")]),
                consumption: numberValue(row[indexOf("消费景气指数")]),
                investment: numberValue(row[indexOf("投资景气指数")]),
                production: numberValue(row[indexOf("生产景气指数")])
            };

            if (
                item.composite !== null &&
                item.consumption !== null &&
                item.investment !== null &&
                item.production !== null
            ) {
                latest = item;
            }
        }

        if (!latest) {
            throw new Error("ICHI Excel 没有有效最新数据");
        }

        setMetric("ichi-composite", latest.composite);
        setMetric("ichi-consumption", latest.consumption);
        setMetric("ichi-investment", latest.investment);
        setMetric("ichi-production", latest.production);

        setText("ichi-composite-note", `最新：${latest.weeknum}`);
        setText(
            "ichi-consumption-note",
            `距 100 ${latest.consumption >= 100 ? "+" : ""}${(
                latest.consumption - 100
            ).toFixed(2)}`
        );
        setText(
            "ichi-investment-note",
            `距 100 ${latest.investment >= 100 ? "+" : ""}${(
                latest.investment - 100
            ).toFixed(2)}`
        );
        setText(
            "ichi-production-note",
            `距 100 ${latest.production >= 100 ? "+" : ""}${(
                latest.production - 100
            ).toFixed(2)}`
        );
    }

    /* ---------- Fiscal homepage snapshot ---------- */

    function tableObjects(workbook, sheetName, requiredHeaders) {
        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            throw new Error(`财政 Excel 缺少 Sheet：${sheetName}`);
        }

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: true,
            defval: null,
            blankrows: false
        });

        const headerRow = findHeader(rows, requiredHeaders);

        if (headerRow < 0) {
            throw new Error(`财政 Excel 未找到表头：${requiredHeaders.join("、")}`);
        }

        const headers = (rows[headerRow] || []).map(value =>
            String(value ?? "").trim()
        );

        return rows
            .slice(headerRow + 1)
            .filter(row => (row || []).some(value => value !== null && value !== ""))
            .map(row => {
                const object = {};

                headers.forEach((header, index) => {
                    if (header) {
                        object[header] = row[index] ?? null;
                    }
                });

                return object;
            });
    }

    async function loadFiscalHomepage() {
        const workbook = await fetchWorkbook(FISCAL_EXCEL_URL);

        const budgetRows = tableObjects(
            workbook,
            "01_预算口径脉冲",
            ["年份", "广义财政安排 (亿元)", "预算口径脉冲 (pp)"]
        )
            .map(row => ({
                year: numberValue(row["年份"]),
                total: numberValue(row["广义财政安排 (亿元)"]),
                impulse: numberValue(row["预算口径脉冲 (pp)"])
            }))
            .filter(row => row.year !== null)
            .sort((a, b) => a.year - b.year);

        const fiscalRows = tableObjects(
            workbook,
            "02_收支口径脉冲",
            [
                "日期",
                "净脉冲 (pp)",
                "分子效应 (pp)",
                "分母效应 (pp)",
                "收支缺口率 (%)"
            ]
        )
            .map(row => ({
                month: excelMonth(row["日期"]),
                impulse: numberValue(row["净脉冲 (pp)"]),
                numerator: numberValue(row["分子效应 (pp)"]),
                denominator: numberValue(row["分母效应 (pp)"]),
                gapRatio: numberValue(row["收支缺口率 (%)"])
            }))
            .filter(row => row.month);

        const debtRows = tableObjects(
            workbook,
            "03_财政收支缺口与新增债务融资",
            ["日期", "新增债务口径 (%)"]
        )
            .map(row => ({
                month: excelMonth(row["日期"]),
                ratio: numberValue(row["新增债务口径 (%)"])
            }))
            .filter(row => row.month && row.ratio !== null);

        const latestBudget = budgetRows.at(-1);
        const previousBudget = budgetRows.at(-2);
        const latestFiscal = [...fiscalRows]
            .reverse()
            .find(row =>
                row.impulse !== null ||
                row.gapRatio !== null
            );
        const latestDebt = debtRows.at(-1);

        if (latestBudget) {
            setMetric(
                "fiscal-budget-pulse",
                latestBudget.impulse,
                "pp"
            );

            setText(
                "fiscal-budget-note",
                latestBudget.total !== null && previousBudget?.total !== null
                    ? `广义财政安排 ${(latestBudget.total / 10000).toFixed(2)}万亿 ← ${(previousBudget.total / 10000).toFixed(2)}万亿`
                    : "--"
            );
        }

        if (latestFiscal) {
            setMetric(
                "fiscal-balance-pulse",
                latestFiscal.impulse,
                "pp"
            );

            setText(
                "fiscal-balance-note",
                `分子 ${latestFiscal.numerator === null ? "--" : latestFiscal.numerator.toFixed(2)} · 分母 ${latestFiscal.denominator === null ? "--" : latestFiscal.denominator.toFixed(2)}`
            );

            setMetric(
                "fiscal-gap-ratio",
                latestFiscal.gapRatio,
                "%"
            );

            const previousYearMonth = latestFiscal.month
                ? `${Number(latestFiscal.month.slice(0, 4)) - 1}-${latestFiscal.month.slice(5, 7)}`
                : null;

            const previousFiscal = fiscalRows.find(
                row => row.month === previousYearMonth
            );

            setText(
                "fiscal-gap-note",
                previousFiscal?.gapRatio !== null &&
                previousFiscal?.gapRatio !== undefined
                    ? `上年同期 ${previousFiscal.gapRatio.toFixed(2)}%`
                    : "--"
            );
        }

        if (latestDebt) {
            const previousYearMonth =
                `${Number(latestDebt.month.slice(0, 4)) - 1}-${latestDebt.month.slice(5, 7)}`;

            const previousDebt = debtRows.find(
                row => row.month === previousYearMonth
            );

            const debtPulse =
                previousDebt?.ratio !== null &&
                previousDebt?.ratio !== undefined
                    ? latestDebt.ratio - previousDebt.ratio
                    : null;

            setMetric(
                "fiscal-debt-pulse",
                debtPulse,
                "pp"
            );

            setText(
                "fiscal-debt-note",
                previousDebt
                    ? `新增债务口径 ${latestDebt.ratio.toFixed(2)}% · 上年同期 ${previousDebt.ratio.toFixed(2)}%`
                    : `新增债务口径 ${latestDebt.ratio.toFixed(2)}%`
            );
        }
    }

    /* ---------- Load dashboard data ---------- */

    async function loadHomepageData() {
        if (typeof XLSX === "undefined") {
            console.warn("SheetJS 未加载，首页显示占位符。");
            return;
        }

        const results = await Promise.allSettled([
            loadICHIHomepage(),
            loadFiscalHomepage()
        ]);

        results.forEach(result => {
            if (result.status === "rejected") {
                console.warn("首页数据读取失败：", result.reason);
            }
        });
    }

    loadHomepageData();
})();

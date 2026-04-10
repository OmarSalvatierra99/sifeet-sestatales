(function () {
  const yearChips = document.getElementById("comparisonYearChips");
  const selectedYearsLabel = document.getElementById("comparisonSelectedYearsLabel");
  const universeLabel = document.getElementById("comparisonUniverseLabel");
  const statusLabel = document.getElementById("comparisonStatus");
  const totalsInsight = document.getElementById("comparisonTotalsInsight");
  const statusInsight = document.getElementById("comparisonStatusInsight");
  const anexoInsight = document.getElementById("comparisonAnexoInsight");
  const pdpInsight = document.getElementById("comparisonPdpInsight");
  const variationInsight = document.getElementById("comparisonVariationInsight");
  const topEntitiesInsight = document.getElementById("comparisonTopEntitiesInsight");
  const dirtyNotice = document.getElementById("comparisonDirtyNotice");
  const universeSelect = document.getElementById("comparison_universe");
  const enteSelect = document.getElementById("comparison_ente_uid");
  const tipoAuditoriaSelect = document.getElementById("comparison_tipo_auditoria");
  const tipoAnexoSelect = document.getElementById("comparison_tipo_anexo");
  const estadoSelect = document.getElementById("comparison_estado");
  const fuenteSelect = document.getElementById("comparison_fuente");
  const origenFuenteSelect = document.getElementById("comparison_origen_fuente");
  const ramoSelect = document.getElementById("comparison_ramo");
  const resetButton = document.getElementById("comparisonResetBtn");
  const clearFab = document.getElementById("comparison_clear_fab");
  const kpiGrid = document.getElementById("comparisonKpiGrid");
  const totalsCanvas = document.getElementById("comparisonTotalsChart");
  const statusYearCanvas = document.getElementById("comparisonStatusYearChart");
  const variationCanvas = document.getElementById("comparisonVariationChart");
  const anexoTotalsCanvas = document.getElementById("comparisonAnexoTotalsChart");
  const stackedCanvas = document.getElementById("comparisonStackedChart");
  const pdpCanvas = document.getElementById("comparisonPdpChart");
  const topChangesList = document.getElementById("comparisonTopChangesList");
  const tableHead = document.getElementById("comparisonTableHead");
  const tableBody = document.getElementById("comparisonTableBody");
  const tableCount = document.getElementById("comparisonTableCount");
  const tableSearch = document.getElementById("comparisonTableSearch");

  if (!yearChips || !window.Chart) {
    return;
  }

  const numberFormatter = new Intl.NumberFormat("es-MX");
  const currencyFormatter = new Intl.NumberFormat("es-MX", {
    style: "currency",
    currency: "MXN",
    maximumFractionDigits: 0,
  });
  const percentFormatter = new Intl.NumberFormat("es-MX", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 1,
  });
  const semanticColors = {
    neutral: "#64748b",
    neutralSoft: "rgba(100, 116, 139, 0.18)",
    success: "#004F61",
    successSoft: "rgba(0, 79, 97, 0.18)",
    danger: "#b91c1c",
    dangerSoft: "rgba(185, 28, 28, 0.18)",
  };
  const technicalPalette = [
    "rgba(0, 79, 97, 0.84)",
    "rgba(0, 105, 126, 0.8)",
    "rgba(8, 126, 164, 0.76)",
    "rgba(202, 138, 4, 0.78)",
    "rgba(127, 29, 29, 0.74)",
    "rgba(71, 85, 105, 0.74)",
  ];

  const createDefaultFilters = () => ({
    universo: "all",
    ente_uid: "",
    tipo_auditoria: "",
    tipo_anexo: "",
    estado: "",
    fuente_financiamiento: "",
    origen_fuente: "",
    ramo_33: "",
  });

  const state = {
    availableYears: [],
    selectedYears: [],
    filters: createDefaultFilters(),
    filterOptions: {
      entes: [],
      tipo_auditoria: [],
      tipo_anexo: [],
      estado: [],
      fuente_financiamiento: [],
      origen_fuente: [],
      ramo_33: [],
    },
    summary: null,
    loading: false,
  };

  const draft = {
    selectedYears: [],
    filters: createDefaultFilters(),
  };

  const charts = {
    totals: null,
    statusYear: null,
    variation: null,
    anexoTotals: null,
    stacked: null,
    pdp: null,
  };

  let requestController = null;

  const escapeHtml = (value) => String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

  const formatNumber = (value) => numberFormatter.format(Number(value || 0));
  const formatCurrency = (value) => currencyFormatter.format(Number(value || 0));
  const formatPercent = (value) => `${percentFormatter.format(Number(value || 0))}%`;
  const formatPercentagePointDelta = (value) => {
    const amount = Number(value || 0);
    const prefix = amount > 0 ? "+" : "";
    return `${prefix}${percentFormatter.format(amount)} pp`;
  };

  const formatDeltaAbs = (value) => {
    const amount = Number(value || 0);
    if (amount > 0) {
      return `+${formatNumber(amount)}`;
    }
    return formatNumber(amount);
  };

  const compactCurrencyTick = (value) => {
    const amount = Number(value || 0);
    if (Math.abs(amount) >= 1000000) {
      return `$${(amount / 1000000).toFixed(1)}M`;
    }
    if (Math.abs(amount) >= 1000) {
      return `$${(amount / 1000).toFixed(1)}k`;
    }
    return `$${formatNumber(amount)}`;
  };

  const shortenLabel = (value, maxLength = 36) => {
    const text = String(value || "").trim();
    if (text.length <= maxLength) {
      return text;
    }
    return `${text.slice(0, maxLength - 1)}...`;
  };

  const cloneFilters = (filters) => ({
    universo: filters.universo || "all",
    ente_uid: filters.ente_uid || "",
    tipo_auditoria: filters.tipo_auditoria || "",
    tipo_anexo: filters.tipo_anexo || "",
    estado: filters.estado || "",
    fuente_financiamiento: filters.fuente_financiamiento || "",
    origen_fuente: filters.origen_fuente || "",
    ramo_33: filters.ramo_33 || "",
  });

  const syncDraftWithState = () => {
    draft.selectedYears = Array.isArray(state.selectedYears) ? [...state.selectedYears] : [];
    draft.filters = cloneFilters(state.filters || {});
  };

  const getChangeClass = (changeLabel, deltaAbs) => {
    if (changeLabel === "Subio" || Number(deltaAbs || 0) > 0) {
      return "is-positive";
    }
    if (changeLabel === "Bajo" || Number(deltaAbs || 0) < 0) {
      return "is-negative";
    }
    return "is-neutral";
  };

  const destroyChart = (key) => {
    if (charts[key]) {
      charts[key].destroy();
      charts[key] = null;
    }
  };

  const setLoading = (loading) => {
    state.loading = loading;
    document.querySelectorAll(".comparativo-page .panel").forEach((panel) => {
      panel.classList.toggle("is-loading", loading);
    });
    [resetButton, clearFab].forEach((element) => {
      if (element) {
        element.disabled = loading;
      }
    });
  };

  const buildParamsFromDraft = () => {
    const params = new URLSearchParams();
    draft.selectedYears.forEach((year) => params.append("anios", year));
    if (draft.filters.universo) {
      params.set("universo", draft.filters.universo);
    }
    if (draft.filters.ente_uid) {
      params.append("ente_uid", draft.filters.ente_uid);
    }
    if (draft.filters.tipo_auditoria) {
      params.append("tipo_auditoria", draft.filters.tipo_auditoria);
    }
    if (draft.filters.tipo_anexo) {
      params.append("tipo_anexo", draft.filters.tipo_anexo);
    }
    if (draft.filters.estado) {
      params.append("estado", draft.filters.estado);
    }
    if (draft.filters.fuente_financiamiento) {
      params.append("fuente_financiamiento", draft.filters.fuente_financiamiento);
    }
    if (draft.filters.origen_fuente) {
      params.append("origen_fuente", draft.filters.origen_fuente);
    }
    if (draft.filters.ramo_33) {
      params.append("ramo_33", draft.filters.ramo_33);
    }
    return params.toString();
  };

  const syncSelectOptions = (selectEl, items, selectedValue, itemToOption) => {
    if (!selectEl) {
      return;
    }
    const currentValue = selectedValue || "";
    const optionList = Array.isArray(items) ? items : [];
    const defaultOption = selectEl.querySelector("option[value='']");
    selectEl.innerHTML = "";
    if (defaultOption) {
      selectEl.appendChild(defaultOption.cloneNode(true));
    }
    optionList.forEach((item) => {
      const option = document.createElement("option");
      const normalized = itemToOption ? itemToOption(item) : { value: item, label: item };
      option.value = normalized.value;
      option.textContent = normalized.label;
      selectEl.appendChild(option);
    });
    if (currentValue && Array.from(selectEl.options).some((option) => option.value === currentValue)) {
      selectEl.value = currentValue;
    } else {
      selectEl.value = "";
    }
  };

  const renderYearChips = () => {
    if (!state.availableYears.length) {
      yearChips.innerHTML = '<span class="muted">No hay ejercicios disponibles.</span>';
      return;
    }
    yearChips.innerHTML = state.availableYears.map((year) => {
      const isActive = draft.selectedYears.includes(year);
      return `
        <button
          class="comparison-year-chip${isActive ? " is-active" : ""}"
          type="button"
          data-year="${escapeHtml(year)}"
          aria-pressed="${isActive ? "true" : "false"}"
        >
          ${escapeHtml(year)}
        </button>
      `;
    }).join("");
  };

  const sumRowCounts = (row) => state.selectedYears.reduce(
    (acc, year) => acc + Number(((row || {}).counts_by_year || {})[year] || 0),
    0,
  );

  const renderSelectionPreview = () => {
    const activeYears = Array.isArray(draft.selectedYears) && draft.selectedYears.length
      ? draft.selectedYears
      : state.selectedYears;
    if (selectedYearsLabel) {
      selectedYearsLabel.textContent = activeYears.length ? activeYears.join(" · ") : "Sin años";
    }
    if (universeLabel) {
      universeLabel.textContent = draft.filters.universo === "complete"
        ? "Solo entes presentes en todos los años"
        : "Universo completo";
    }
    if (dirtyNotice) {
      dirtyNotice.textContent = "Los filtros se aplican automaticamente al seleccionar.";
      dirtyNotice.classList.remove("is-dirty");
    }
  };

  const buildTotalsInsight = () => {
    const rows = (state.summary && state.summary.totals_by_year) || [];
    if (!rows.length) {
      return "No hay observaciones para los filtros aplicados.";
    }
    if (rows.length === 1) {
      return `En ${rows[0].ejercicio} se registraron ${formatNumber(rows[0].total_observaciones)} observaciones.`;
    }
    const first = rows[0];
    const last = rows[rows.length - 1];
    const delta = Number(last.total_observaciones || 0) - Number(first.total_observaciones || 0);
    if (delta > 0) {
      return `En ${last.ejercicio} hubo ${formatNumber(delta)} observaciones mas que en ${first.ejercicio}.`;
    }
    if (delta < 0) {
      return `En ${last.ejercicio} hubo ${formatNumber(Math.abs(delta))} observaciones menos que en ${first.ejercicio}.`;
    }
    return `En ${last.ejercicio} se mantuvo el mismo total de observaciones que en ${first.ejercicio}.`;
  };

  const buildStatusInsight = () => {
    const rows = (state.summary && state.summary.kpis_by_year) || [];
    if (!rows.length) {
      return "No hay informacion de solventacion para mostrar.";
    }
    const first = rows[0];
    const latest = rows[rows.length - 1];
    const latestRate = Number(latest.porcentaje_solventacion || 0);
    if (rows.length === 1) {
      return `En ${latest.ejercicio}, la tasa de solventacion fue ${formatPercent(latestRate)}.`;
    }
    const firstRate = Number(first.porcentaje_solventacion || 0);
    const delta = latestRate - firstRate;
    if (Math.abs(delta) < 0.05) {
      return `La tasa de solventacion se mantuvo en ${formatPercent(latestRate)} entre ${first.ejercicio} y ${latest.ejercicio}.`;
    }
    return `En ${latest.ejercicio}, la tasa de solventacion fue ${formatPercent(latestRate)} (${delta > 0 ? "subio" : "bajo"} ${formatPercentagePointDelta(delta)} frente a ${first.ejercicio}).`;
  };

  const buildVariationInsight = () => {
    const topChanges = (((state.summary || {}).top_variations || {}).top_changes) || [];
    if (state.selectedYears.length < 2) {
      return "Selecciona al menos dos años para calcular cambios.";
    }
    if (!topChanges.length) {
      return "No hubo cambios entre los anos seleccionados.";
    }
    const first = topChanges[0];
    return `El mayor cambio fue ${first.label} con ${formatDeltaAbs(first.delta_abs)} observaciones.`;
  };

  const buildAnexoInsight = () => {
    const rows = (state.summary && state.summary.anexo_totals_by_year) || [];
    if (!rows.length) {
      return "No hay tipos de anexo para mostrar.";
    }
    const latestYear = state.selectedYears[state.selectedYears.length - 1] || "";
    const latestRows = rows.filter((row) => row.ejercicio === latestYear);
    if (!latestRows.length) {
      return "No hay tipos de anexo para mostrar.";
    }
    const topRow = latestRows.slice().sort((left, right) => Number(right.total || 0) - Number(left.total || 0))[0];
    const totalLatest = latestRows.reduce((acc, row) => acc + Number(row.total || 0), 0);
    const share = totalLatest ? (Number(topRow.total || 0) / totalLatest) * 100 : 0;
    return `En ${latestYear}, ${topRow.tipo_anexo} concentro ${formatPercent(share)} del total observado.`;
  };

  const buildPdpInsight = () => {
    const rows = (state.summary && state.summary.pdp_amounts_by_year) || [];
    if (!rows.length) {
      return "No hay montos PDP para mostrar.";
    }
    const first = rows[0];
    const latest = rows[rows.length - 1];
    if (rows.length === 1) {
      return `En ${latest.ejercicio}, el monto PDP pendiente fue ${formatCurrency(latest.pendiente || 0)}.`;
    }
    const delta = Number(latest.pendiente || 0) - Number(first.pendiente || 0);
    if (delta === 0) {
      return `El monto PDP pendiente se mantuvo en ${formatCurrency(latest.pendiente || 0)} entre ${first.ejercicio} y ${latest.ejercicio}.`;
    }
    return `El monto PDP pendiente en ${latest.ejercicio} fue ${formatCurrency(latest.pendiente || 0)} (${delta > 0 ? "subio" : "bajo"} ${formatCurrency(Math.abs(delta))} frente a ${first.ejercicio}).`;
  };

  const buildTopEntitiesInsight = () => {
    const rows = Array.isArray((state.summary || {}).comparison_table) ? state.summary.comparison_table : [];
    if (!rows.length) {
      return "No hay entes para mostrar.";
    }
    const ordered = rows.slice().sort((left, right) => sumRowCounts(right) - sumRowCounts(left));
    const topRow = ordered[0];
    return `${topRow.label} concentro ${formatNumber(sumRowCounts(topRow))} observaciones acumuladas en los anos seleccionados.`;
  };

  const renderMeta = () => {
    const totals = (state.summary && state.summary.totals_by_year) || [];
    renderSelectionPreview();
    if (universeLabel) {
      universeLabel.textContent = ((state.summary || {}).universo || {}).mode_label || "Universo completo";
    }
    if (totalsInsight) {
      totalsInsight.textContent = buildTotalsInsight();
    }
    if (statusInsight) {
      statusInsight.textContent = buildStatusInsight();
    }
    if (anexoInsight) {
      anexoInsight.textContent = buildAnexoInsight();
    }
    if (pdpInsight) {
      pdpInsight.textContent = buildPdpInsight();
    }
    if (variationInsight) {
      variationInsight.textContent = buildVariationInsight();
    }
    if (topEntitiesInsight) {
      topEntitiesInsight.textContent = buildTopEntitiesInsight();
    }

    if (statusLabel) {
      statusLabel.textContent = totals.length ? "Filtros sincronizados" : "Sin resultados";
    }
  };

  const renderFilters = () => {
    const options = state.filterOptions || {};
    if (universeSelect) {
      universeSelect.value = draft.filters.universo || "all";
    }
    syncSelectOptions(
      enteSelect,
      options.entes || [],
      draft.filters.ente_uid,
      (item) => {
        const suffix = item.has_historical_names ? " · nombre historico" : "";
        return { value: item.ente_uid, label: `${item.label}${suffix}` };
      },
    );
    syncSelectOptions(tipoAuditoriaSelect, options.tipo_auditoria || [], draft.filters.tipo_auditoria);
    syncSelectOptions(tipoAnexoSelect, options.tipo_anexo || [], draft.filters.tipo_anexo);
    syncSelectOptions(estadoSelect, options.estado || [], draft.filters.estado);
    syncSelectOptions(
      fuenteSelect,
      options.fuente_financiamiento || [],
      draft.filters.fuente_financiamiento,
    );
    syncSelectOptions(origenFuenteSelect, options.origen_fuente || [], draft.filters.origen_fuente);
    syncSelectOptions(ramoSelect, options.ramo_33 || [], draft.filters.ramo_33);
  };

  const renderKpis = () => {
    const rows = (state.summary && state.summary.kpis_by_year) || [];
    if (!rows.length) {
      kpiGrid.innerHTML = `
        <article class="comparison-empty-card">
          <p>No hay observaciones para los filtros seleccionados.</p>
        </article>
      `;
      return;
    }

    kpiGrid.innerHTML = rows.map((row) => {
      return `
        <article class="comparison-kpi-card">
          <div class="comparison-kpi-year">${escapeHtml(row.ejercicio)}</div>
          <dl class="comparison-kpi-list">
            <div>
              <dt>Emitidas</dt>
              <dd>${formatNumber(row.emitidas)}</dd>
            </div>
            <div>
              <dt>Pendientes</dt>
              <dd>${formatNumber(row.pendientes)}</dd>
            </div>
            <div>
              <dt>Solventadas</dt>
              <dd>${formatNumber(row.solventadas)}</dd>
            </div>
            <div>
              <dt>Solventación</dt>
              <dd>${formatPercent(row.porcentaje_solventacion || 0)}</dd>
            </div>
          </dl>
        </article>
      `;
    }).join("");
  };

  const renderTotalsChart = () => {
    destroyChart("totals");
    if (!totalsCanvas) {
      return;
    }
    const rows = (state.summary && state.summary.totals_by_year) || [];
    const data = rows.map((item) => Number(item.total_observaciones || 0));
    charts.totals = new Chart(totalsCanvas, {
      type: "line",
      data: {
        labels: rows.map((item) => item.ejercicio),
        datasets: [
          {
            label: "Observaciones emitidas",
            data,
            backgroundColor: semanticColors.neutralSoft,
            borderColor: semanticColors.neutral,
            borderWidth: 2,
            fill: true,
            tension: 0.28,
            pointRadius: 4,
            pointHoverRadius: 5,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              precision: 0,
              callback: (value) => formatNumber(value),
            },
          },
        },
      },
    });
  };

  const renderStatusYearChart = () => {
    destroyChart("statusYear");
    if (!statusYearCanvas) {
      return;
    }
    const rows = (state.summary && state.summary.kpis_by_year) || [];
    charts.statusYear = new Chart(statusYearCanvas, {
      type: "line",
      data: {
        labels: rows.map((item) => item.ejercicio),
        datasets: [
          {
            label: "Tasa de solventacion",
            data: rows.map((item) => Number(item.porcentaje_solventacion || 0)),
            backgroundColor: semanticColors.successSoft,
            borderColor: semanticColors.success,
            borderWidth: 2,
            fill: true,
            tension: 0.28,
            pointRadius: 4,
            pointHoverRadius: 5,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false,
          },
          tooltip: {
            callbacks: {
              label: (context) => `Solventacion: ${formatPercent(context.parsed.y)}`,
            },
          },
        },
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            ticks: {
              callback: (value) => formatPercent(value),
            },
          },
        },
      },
    });
  };

  const renderVariationChart = () => {
    destroyChart("variation");
    if (!variationCanvas) {
      return;
    }
    const topChanges = ((((state.summary || {}).top_variations || {}).top_changes) || []).slice(0, 5);
    charts.variation = new Chart(variationCanvas, {
      type: "bar",
      data: {
        labels: topChanges.map((item) => shortenLabel(item.label)),
        datasets: [
          {
            label: "Cambio absoluto",
            data: topChanges.map((item) => Number(item.delta_abs || 0)),
            backgroundColor: topChanges.map((item) => (
              Number(item.delta_abs || 0) >= 0
                ? semanticColors.successSoft
                : semanticColors.dangerSoft
            )),
            borderColor: topChanges.map((item) => (
              Number(item.delta_abs || 0) >= 0
                ? semanticColors.success
                : semanticColors.danger
            )),
            borderWidth: 1.2,
            borderRadius: 10,
          },
        ],
      },
      options: {
        indexAxis: "y",
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              title: (items) => {
                const index = items && items[0] ? items[0].dataIndex : -1;
                return index >= 0 && topChanges[index] ? topChanges[index].label : "";
              },
              label: (context) => `Cambio absoluto: ${formatDeltaAbs(context.parsed.x)}`,
            },
          },
        },
        scales: {
          x: {
            ticks: {
              precision: 0,
              callback: (value) => formatNumber(value),
            },
          },
        },
      },
    });
  };

  const collapseTopAnexoRows = (rows) => {
    const totalsByType = new Map();
    (Array.isArray(rows) ? rows : []).forEach((row) => {
      const tipo = String(row.tipo_anexo || "").trim();
      if (!tipo) {
        return;
      }
      const bucketTotal = Number(row.total || 0) + Number(row.pendientes || 0) + Number(row.solventadas || 0);
      totalsByType.set(tipo, (totalsByType.get(tipo) || 0) + bucketTotal);
    });
    const orderedTypes = Array.from(totalsByType.entries())
      .sort((left, right) => right[1] - left[1])
      .map((item) => item[0]);
    const topTypes = orderedTypes.slice(0, 4);
    const useOther = orderedTypes.length > 4;
    return {
      labels: useOther ? [...topTypes, "Otros"] : topTypes,
      bucketFor: (tipo) => (topTypes.includes(tipo) ? tipo : "Otros"),
    };
  };

  const renderAnexoTotalsChart = () => {
    destroyChart("anexoTotals");
    if (!anexoTotalsCanvas) {
      return;
    }
    const rows = (state.summary && state.summary.anexo_totals_by_year) || [];
    const collapsed = collapseTopAnexoRows(rows);
    const bucketMap = new Map();
    const totalByYear = new Map();
    rows.forEach((row) => {
      const year = row.ejercicio;
      const total = Number(row.total || 0);
      const bucket = collapsed.bucketFor(String(row.tipo_anexo || "").trim());
      const key = `${year}||${bucket}`;
      bucketMap.set(key, (bucketMap.get(key) || 0) + total);
      totalByYear.set(year, (totalByYear.get(year) || 0) + total);
    });
    charts.anexoTotals = new Chart(anexoTotalsCanvas, {
      type: "bar",
      data: {
        labels: state.selectedYears,
        datasets: collapsed.labels.map((bucket, index) => {
          const rawCounts = state.selectedYears.map((year) => Number(bucketMap.get(`${year}||${bucket}`) || 0));
          return {
            label: bucket,
            data: state.selectedYears.map((year, dataIndex) => {
              const yearTotal = Number(totalByYear.get(year) || 0);
              const count = rawCounts[dataIndex] || 0;
              return yearTotal ? Number(((count / yearTotal) * 100).toFixed(1)) : 0;
            }),
            rawCounts,
            backgroundColor: technicalPalette[index % technicalPalette.length],
            borderRadius: 8,
          };
        }),
      },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "bottom",
          },
          tooltip: {
            callbacks: {
              label: (context) => {
                const rawCounts = Array.isArray(context.dataset.rawCounts) ? context.dataset.rawCounts : [];
                const count = Number(rawCounts[context.dataIndex] || 0);
                return `${context.dataset.label}: ${formatPercent(context.parsed.y)} (${formatNumber(count)} observaciones)`;
              },
            },
          },
        },
        scales: {
          x: {
            stacked: true,
          },
          y: {
            stacked: true,
            beginAtZero: true,
            max: 100,
            ticks: {
              callback: (value) => formatPercent(value),
            },
          },
        },
      },
    });
  };

  const renderStackedChart = () => {
    destroyChart("stacked");
    if (!stackedCanvas) {
      return;
    }
    const rows = Array.isArray((state.summary || {}).comparison_table) ? state.summary.comparison_table : [];
    const topRows = rows.slice()
      .sort((left, right) => sumRowCounts(right) - sumRowCounts(left))
      .slice(0, 5);
    charts.stacked = new Chart(stackedCanvas, {
      type: "bar",
      data: {
        labels: topRows.map((row) => shortenLabel(row.label)),
        datasets: state.selectedYears.map((year, index) => ({
          label: year,
          data: topRows.map((row) => Number((row.counts_by_year || {})[year] || 0)),
          backgroundColor: technicalPalette[index % technicalPalette.length],
          borderRadius: 8,
        })),
      },
      options: {
        indexAxis: "y",
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "bottom",
          },
          tooltip: {
            callbacks: {
              title: (items) => {
                const index = items && items[0] ? items[0].dataIndex : -1;
                return index >= 0 && topRows[index] ? topRows[index].label : "";
              },
              label: (context) => `${context.dataset.label}: ${formatNumber(context.parsed.x)}`,
            },
          },
        },
        scales: {
          x: {
            beginAtZero: true,
            ticks: {
              precision: 0,
              callback: (value) => formatNumber(value),
            },
          },
        },
      },
    });
  };

  const renderPdpChart = () => {
    destroyChart("pdp");
    if (!pdpCanvas) {
      return;
    }
    const rows = (state.summary && state.summary.pdp_amounts_by_year) || [];
    charts.pdp = new Chart(pdpCanvas, {
      type: "bar",
      data: {
        labels: rows.map((item) => item.ejercicio),
        datasets: [
          {
            label: "Emitido",
            data: rows.map((item) => Number(item.emitido || 0)),
            backgroundColor: semanticColors.neutralSoft,
            borderColor: semanticColors.neutral,
            borderWidth: 1.2,
            borderRadius: 8,
          },
          {
            label: "Solventado",
            data: rows.map((item) => Number(item.solventado || 0)),
            backgroundColor: semanticColors.successSoft,
            borderColor: semanticColors.success,
            borderWidth: 1.2,
            borderRadius: 8,
          },
          {
            label: "Pendiente",
            data: rows.map((item) => Number(item.pendiente || 0)),
            backgroundColor: semanticColors.dangerSoft,
            borderColor: semanticColors.danger,
            borderWidth: 1.2,
            borderRadius: 8,
          },
        ],
      },
      options: {
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "bottom",
          },
          tooltip: {
            callbacks: {
              label: (context) => `${context.dataset.label}: ${formatCurrency(context.parsed.y)}`,
            },
          },
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              callback: (value) => compactCurrencyTick(value),
            },
          },
        },
      },
    });
  };

  const renderTopChangesList = () => {
    if (!topChangesList) {
      return;
    }
    const topChanges = ((((state.summary || {}).top_variations || {}).top_changes) || []).slice(0, 5);
    if (state.selectedYears.length < 2) {
      topChangesList.innerHTML = '<p class="muted">Selecciona al menos dos anos para calcular variaciones.</p>';
      return;
    }
    if (!topChanges.length) {
      topChangesList.innerHTML = '<p class="muted">No hubo cambios para los filtros seleccionados.</p>';
      return;
    }
    const firstYear = state.selectedYears[0] || "";
    const lastYear = state.selectedYears[state.selectedYears.length - 1] || "";
    topChangesList.innerHTML = topChanges.map((item) => {
      const firstValue = Number((item.counts_by_year || {})[firstYear] || 0);
      const lastValue = Number((item.counts_by_year || {})[lastYear] || 0);
      const changeClass = getChangeClass(item.change_label, item.delta_abs);
      return `
        <article class="comparison-top-change-item">
          <div class="comparison-top-change-copy">
            <strong>${escapeHtml(item.label)}</strong>
            <span>${escapeHtml(firstYear)}: ${formatNumber(firstValue)} · ${escapeHtml(lastYear)}: ${formatNumber(lastValue)}</span>
          </div>
          <div class="comparison-top-change-meta">
            <span class="comparison-delta-badge ${changeClass}">${escapeHtml(formatDeltaAbs(item.delta_abs))}</span>
            <span class="comparison-change-label">${escapeHtml(item.change_label || "Sin cambio")}</span>
          </div>
        </article>
      `;
    }).join("");
  };

  const renderTable = () => {
    const summary = state.summary || {};
    const rows = Array.isArray(summary.comparison_table) ? summary.comparison_table : [];
    const searchTerm = String(tableSearch ? tableSearch.value || "" : "").trim().toLowerCase();
    const filteredRows = rows.filter((row) => {
      if (!searchTerm) {
        return true;
      }
      const aliases = Array.isArray(row.aliases) ? row.aliases.join(" ") : "";
      const haystack = `${row.label || ""} ${row.ente_uid || ""} ${aliases}`.toLowerCase();
      return haystack.includes(searchTerm);
    });

    tableHead.innerHTML = `
      <tr>
        <th>Ente</th>
        ${state.selectedYears.map((year) => `<th>${escapeHtml(year)}</th>`).join("")}
        <th>Cambio abs.</th>
        <th>Lectura</th>
      </tr>
    `;

    if (!filteredRows.length) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="${state.selectedYears.length + 3}" class="muted">Sin resultados para la busqueda o los filtros activos.</td>
        </tr>
      `;
      tableCount.textContent = "0 entes visibles.";
      return;
    }

    tableBody.innerHTML = filteredRows.map((row) => {
      const aliases = Array.isArray(row.aliases) ? row.aliases.filter(Boolean) : [];
      const aliasLine = row.has_historical_names
        ? `<span class="comparison-table-aliases">${escapeHtml(aliases.join(" · "))}</span>`
        : "";
      const changeClass = getChangeClass(row.change_label, row.delta_abs);
      return `
        <tr>
          <td class="comparison-ente-cell">
            <strong>${escapeHtml(row.label)}</strong>
            ${aliasLine}
          </td>
          ${state.selectedYears.map((year) => {
            const count = row.counts_by_year && Object.prototype.hasOwnProperty.call(row.counts_by_year, year)
              ? row.counts_by_year[year]
              : 0;
            return `<td>${formatNumber(count)}</td>`;
          }).join("")}
          <td><span class="comparison-delta-badge ${changeClass}">${escapeHtml(formatDeltaAbs(row.delta_abs))}</span></td>
          <td><span class="comparison-change-chip ${changeClass}">${escapeHtml(row.change_label || "Sin cambio")}</span></td>
        </tr>
      `;
    }).join("");
    tableCount.textContent = `${formatNumber(filteredRows.length)} de ${formatNumber(rows.length)} entes visibles.`;
  };

  const renderAll = () => {
    renderYearChips();
    renderFilters();
    renderMeta();
    renderKpis();
    renderTotalsChart();
    renderStatusYearChart();
    renderVariationChart();
    renderAnexoTotalsChart();
    renderStackedChart();
    renderPdpChart();
    renderTopChangesList();
    renderTable();
  };

  const loadStats = async () => {
    let shouldRenderMeta = true;
    if (requestController) {
      requestController.abort();
    }
    requestController = new AbortController();
    setLoading(true);
    if (statusLabel) {
      statusLabel.textContent = "Consultando comparativo...";
    }
    try {
      const query = buildParamsFromDraft();
      const response = await fetch(`/comparativo-anual/stats${query ? `?${query}` : ""}`, {
        signal: requestController.signal,
      });
      if (!response.ok) {
        throw new Error("No se pudo consultar el comparativo anual.");
      }
      const data = await response.json();
      state.availableYears = Array.isArray(data.available_years) ? data.available_years : [];
      state.selectedYears = Array.isArray(data.selected_years) ? data.selected_years : [];
      state.filterOptions = data.filter_options || state.filterOptions;
      const selectedFilters = data.selected_filters || {};
      state.filters = {
        universo: selectedFilters.universo || "all",
        ente_uid: Array.isArray(selectedFilters.ente_uid) ? (selectedFilters.ente_uid[0] || "") : "",
        tipo_auditoria: Array.isArray(selectedFilters.tipo_auditoria) ? (selectedFilters.tipo_auditoria[0] || "") : "",
        tipo_anexo: Array.isArray(selectedFilters.tipo_anexo) ? (selectedFilters.tipo_anexo[0] || "") : "",
        estado: Array.isArray(selectedFilters.estado) ? (selectedFilters.estado[0] || "") : "",
        fuente_financiamiento: Array.isArray(selectedFilters.fuente_financiamiento)
          ? (selectedFilters.fuente_financiamiento[0] || "")
          : "",
        origen_fuente: Array.isArray(selectedFilters.origen_fuente)
          ? (selectedFilters.origen_fuente[0] || "")
          : "",
        ramo_33: Array.isArray(selectedFilters.ramo_33) ? (selectedFilters.ramo_33[0] || "") : "",
      };
      state.summary = data.summary || {};
      syncDraftWithState();
      renderAll();
    } catch (error) {
      if (error && error.name === "AbortError") {
        return;
      }
      shouldRenderMeta = false;
      if (statusLabel) {
        statusLabel.textContent = error && error.message
          ? error.message
          : "No se pudo cargar el comparativo anual.";
      }
      kpiGrid.innerHTML = `
        <article class="comparison-empty-card">
          <p>Error al consultar el comparativo.</p>
        </article>
      `;
    } finally {
      setLoading(false);
      if (shouldRenderMeta) {
        renderMeta();
      } else {
        renderSelectionPreview();
      }
    }
  };

  yearChips.addEventListener("click", (event) => {
    const button = event.target.closest("[data-year]");
    if (!button) {
      return;
    }
    const year = button.getAttribute("data-year") || "";
    if (!year) {
      return;
    }
    const selected = new Set(draft.selectedYears);
    if (selected.has(year)) {
      if (selected.size === 1) {
        return;
      }
      selected.delete(year);
    } else {
      selected.add(year);
    }
    draft.selectedYears = state.availableYears.filter((item) => selected.has(item));
    renderYearChips();
    renderSelectionPreview();
    loadStats();
  });

  [
    [universeSelect, "universo"],
    [enteSelect, "ente_uid"],
    [tipoAuditoriaSelect, "tipo_auditoria"],
    [tipoAnexoSelect, "tipo_anexo"],
    [estadoSelect, "estado"],
    [fuenteSelect, "fuente_financiamiento"],
    [origenFuenteSelect, "origen_fuente"],
    [ramoSelect, "ramo_33"],
  ].forEach(([element, key]) => {
    if (!element) {
      return;
    }
    element.addEventListener("change", () => {
      draft.filters[key] = element.value || "";
      renderSelectionPreview();
      loadStats();
    });
  });

  const resetComparisonFilters = () => {
    draft.selectedYears = [];
    draft.filters = createDefaultFilters();
    if (tableSearch) {
      tableSearch.value = "";
    }
    renderSelectionPreview();
    loadStats();
  };

  if (resetButton) {
    resetButton.addEventListener("click", resetComparisonFilters);
  }

  if (clearFab) {
    clearFab.addEventListener("click", resetComparisonFilters);
  }

  if (tableSearch) {
    tableSearch.addEventListener("input", renderTable);
  }

  loadStats();
})();

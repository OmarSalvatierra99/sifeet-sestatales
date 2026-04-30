(function () {
  const states = new WeakMap();

  const escapeSelector = (value) => {
    if (window.CSS && typeof window.CSS.escape === "function") {
      return window.CSS.escape(value);
    }
    return String(value || "").replace(/"/g, '\\"');
  };

  const getConfig = (selectEl) => states.get(selectEl)?.config || {};

  const getSelectedValues = (selectEl) => {
    if (!selectEl) {
      return [];
    }
    if (selectEl.multiple) {
      return Array.from(selectEl.selectedOptions || [])
        .map((option) => option.value)
        .filter(Boolean);
    }
    return selectEl.value ? [selectEl.value] : [];
  };

  const setSelectedValues = (selectEl, values) => {
    if (!selectEl) {
      return;
    }
    const safeValues = Array.isArray(values) ? values.filter(Boolean) : [];
    const valueSet = new Set(safeValues);
    if (selectEl.multiple) {
      Array.from(selectEl.options || []).forEach((option) => {
        option.selected = valueSet.has(option.value);
      });
      return;
    }
    const selectedValue = safeValues.find((value) => Array.from(selectEl.options || []).some((option) => option.value === value));
    selectEl.value = selectedValue || "";
  };

  const getOptionEntries = (selectEl) => Array.from(selectEl?.options || [])
    .filter((option) => option.value)
    .map((option) => ({
      value: option.value,
      label: String(option.textContent || option.value).trim(),
      selected: Boolean(option.selected),
    }));

  const formatSelectedSummary = (selectEl, config) => {
    const placeholder = config.placeholder || "Todos";
    const selected = getSelectedValues(selectEl);
    const total = getOptionEntries(selectEl).length;
    if (!selected.length) {
      return placeholder;
    }
    if (total > 0 && selected.length === total) {
      return config.allSelectedLabel || placeholder;
    }
    if (typeof config.selectedCountLabel === "function") {
      return config.selectedCountLabel(selected.length, total);
    }
    if (typeof config.selectedCountLabel === "string" && config.selectedCountLabel.trim()) {
      return config.selectedCountLabel.includes("{count}")
        ? config.selectedCountLabel.replace("{count}", String(selected.length))
        : `${selected.length} ${config.selectedCountLabel}`;
    }
    return selected.length === 1
      ? "1 seleccionado"
      : `${selected.length} seleccionados`;
  };

  const formatCountLabel = (config, filteredCount, totalCount) => {
    if (typeof config.optionCountLabel === "function") {
      return config.optionCountLabel(filteredCount, totalCount);
    }
    const noun = String(config.optionCountLabel || "opciones").trim();
    if (filteredCount !== totalCount) {
      return `${filteredCount} de ${totalCount} ${noun}`;
    }
    return `${totalCount} ${noun}`;
  };

  const close = (selectEl, { focusToggle = false, clearSearch = true } = {}) => {
    const state = states.get(selectEl);
    if (!state) {
      return;
    }
    state.container.classList.remove("is-open");
    state.container.setAttribute("aria-expanded", "false");
    if (clearSearch) {
      state.searchTerm = "";
    }
    render(selectEl);
    if (focusToggle) {
      window.setTimeout(() => {
        state.toggle.focus();
      }, 0);
    }
  };

  const closeAll = (exceptSelect = null) => {
    states.forEach?.(() => {});
    document.querySelectorAll("[data-multi-select-container].is-open").forEach((container) => {
      const selectId = container.getAttribute("data-select-id");
      if (!selectId) {
        return;
      }
      const selectEl = document.getElementById(selectId);
      if (!selectEl || (exceptSelect && selectEl === exceptSelect)) {
        return;
      }
      close(selectEl);
    });
  };

  const syncDisabled = (selectEl) => {
    const state = states.get(selectEl);
    if (!state) {
      return;
    }
    const disabled = Boolean(selectEl.disabled);
    state.toggle.disabled = disabled;
    state.container.classList.toggle("is-disabled", disabled);
    if (disabled) {
      close(selectEl, { clearSearch: false });
    }
  };

  const focusSearchOrFirstOption = (selectEl) => {
    const state = states.get(selectEl);
    if (!state) {
      return;
    }
    window.setTimeout(() => {
      if (state.searchInput && document.contains(state.searchInput)) {
        state.searchInput.focus();
        const length = state.searchInput.value.length;
        state.searchInput.setSelectionRange(length, length);
        return;
      }
      const firstCheckbox = state.panel.querySelector(".multi-select-checkbox");
      if (firstCheckbox instanceof HTMLElement) {
        firstCheckbox.focus();
      }
    }, 0);
  };

  const toggleOpen = (selectEl, forceOpen = null) => {
    const state = states.get(selectEl);
    if (!state || state.toggle.disabled) {
      return;
    }
    const shouldOpen = forceOpen === null
      ? !state.container.classList.contains("is-open")
      : Boolean(forceOpen);
    if (!shouldOpen) {
      close(selectEl, { clearSearch: true });
      return;
    }
    closeAll(selectEl);
    state.container.classList.add("is-open");
    state.container.setAttribute("aria-expanded", "true");
    render(selectEl);
    focusSearchOrFirstOption(selectEl);
  };

  const moveFocusBetweenCheckboxes = (current, direction) => {
    const checkboxes = Array.from(document.querySelectorAll(".multi-select-checkbox"));
    const index = checkboxes.indexOf(current);
    if (index === -1) {
      return;
    }
    const target = checkboxes[index + direction];
    if (target instanceof HTMLElement) {
      target.focus();
    }
  };

  const render = (selectEl) => {
    const state = states.get(selectEl);
    if (!state) {
      return;
    }
    const config = state.config;
    const options = getOptionEntries(selectEl);
    const searchTerm = String(state.searchTerm || "").trim().toLowerCase();
    const filteredOptions = options.filter((option) => {
      if (!searchTerm) {
        return true;
      }
      const haystack = `${option.label} ${option.value}`.toLowerCase();
      return haystack.includes(searchTerm);
    });

    state.value.textContent = formatSelectedSummary(selectEl, config);
    state.panel.innerHTML = "";

    const panelInner = document.createElement("div");
    panelInner.className = "multi-select-panel-inner";

    if (config.searchable) {
      const searchWrap = document.createElement("div");
      searchWrap.className = "multi-select-search-wrap";
      const searchInput = document.createElement("input");
      searchInput.type = "search";
      searchInput.className = "multi-select-search-input";
      searchInput.placeholder = config.searchPlaceholder || "Buscar";
      searchInput.value = state.searchTerm || "";
      searchInput.setAttribute("aria-label", config.searchPlaceholder || "Buscar");
      searchInput.addEventListener("input", () => {
        state.searchTerm = searchInput.value || "";
        render(selectEl);
      });
      searchInput.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          close(selectEl, { focusToggle: true });
          return;
        }
        if (event.key === "ArrowDown") {
          event.preventDefault();
          const firstCheckbox = state.panel.querySelector(".multi-select-checkbox");
          if (firstCheckbox instanceof HTMLElement) {
            firstCheckbox.focus();
          }
        }
      });
      searchWrap.appendChild(searchInput);
      panelInner.appendChild(searchWrap);
      state.searchInput = searchInput;
    } else {
      state.searchInput = null;
    }

    const toolbar = document.createElement("div");
    toolbar.className = "multi-select-toolbar";

    const countNode = document.createElement("span");
    countNode.className = "multi-select-search-count";
    countNode.textContent = formatCountLabel(config, filteredOptions.length, options.length);
    toolbar.appendChild(countNode);

    const actions = document.createElement("div");
    actions.className = "multi-select-actions";

    if (config.selectAllLabel) {
      const selectAllButton = document.createElement("button");
      selectAllButton.type = "button";
      selectAllButton.className = "multi-select-action-link";
      selectAllButton.textContent = config.selectAllLabel;
      selectAllButton.addEventListener("click", () => {
        options.forEach((option) => {
          const target = Array.from(selectEl.options || []).find((item) => item.value === option.value);
          if (target) {
            target.selected = true;
          }
        });
        selectEl.dispatchEvent(new Event("change", { bubbles: true }));
        render(selectEl);
      });
      actions.appendChild(selectAllButton);
    }

    if (config.clearLabel) {
      const clearButton = document.createElement("button");
      clearButton.type = "button";
      clearButton.className = "multi-select-action-link";
      clearButton.textContent = config.clearLabel;
      clearButton.addEventListener("click", () => {
        Array.from(selectEl.options || []).forEach((option) => {
          option.selected = false;
        });
        selectEl.dispatchEvent(new Event("change", { bubbles: true }));
        render(selectEl);
      });
      actions.appendChild(clearButton);
    }

    if (actions.childElementCount) {
      toolbar.appendChild(actions);
    }
    panelInner.appendChild(toolbar);

    const list = document.createElement("div");
    list.className = "multi-select-options-list";
    list.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        close(selectEl, { focusToggle: true });
      }
    });

    if (!filteredOptions.length) {
      const emptyNode = document.createElement("div");
      emptyNode.className = "multi-select-empty";
      emptyNode.textContent = "Sin resultados para la búsqueda.";
      list.appendChild(emptyNode);
    } else {
      filteredOptions.forEach((option, index) => {
        const row = document.createElement("label");
        row.className = "multi-select-option";
        row.setAttribute("data-option-index", String(index));

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "multi-select-checkbox";
        checkbox.checked = option.selected;
        checkbox.setAttribute("aria-label", option.label);
        checkbox.addEventListener("change", () => {
          const target = Array.from(selectEl.options || []).find((item) => item.value === option.value);
          if (target) {
            target.selected = checkbox.checked;
          }
          selectEl.dispatchEvent(new Event("change", { bubbles: true }));
          render(selectEl);
        });
        checkbox.addEventListener("keydown", (event) => {
          if (event.key === "ArrowDown") {
            event.preventDefault();
            moveFocusBetweenCheckboxes(checkbox, 1);
          } else if (event.key === "ArrowUp") {
            event.preventDefault();
            if (index === 0 && state.searchInput instanceof HTMLElement) {
              state.searchInput.focus();
            } else {
              moveFocusBetweenCheckboxes(checkbox, -1);
            }
          } else if (event.key === "Escape") {
            event.preventDefault();
            close(selectEl, { focusToggle: true });
          }
        });

        const text = document.createElement("span");
        text.className = "multi-select-option-text";
        text.textContent = option.label;

        row.appendChild(checkbox);
        row.appendChild(text);
        list.appendChild(row);
      });
    }

    panelInner.appendChild(list);
    state.panel.appendChild(panelInner);
  };

  const destroy = (selectEl) => {
    const state = states.get(selectEl);
    if (!state) {
      return;
    }
    state.container.remove();
    selectEl.classList.remove("is-multi-hidden");
    states.delete(selectEl);
  };

  const sync = (selectEl, config = {}) => {
    if (!selectEl) {
      return;
    }
    const allowMulti = Boolean(config.multiple);
    if (!allowMulti) {
      destroy(selectEl);
      return;
    }
    const mergedConfig = {
      label: config.label || "",
      placeholder: config.placeholder || "Todos",
      searchable: Boolean(config.searchable),
      searchPlaceholder: config.searchPlaceholder || "",
      selectAllLabel: config.selectAllLabel || "",
      clearLabel: config.clearLabel || "",
      allSelectedLabel: config.allSelectedLabel || config.placeholder || "Todos",
      selectedCountLabel: config.selectedCountLabel || null,
      optionCountLabel: config.optionCountLabel || "opciones",
    };

    let state = states.get(selectEl);
    if (!state) {
      const selectId = selectEl.id || `multi_select_${Math.random().toString(36).slice(2, 10)}`;
      selectEl.id = selectId;
      const container = document.createElement("div");
      container.className = "multi-select-dropdown";
      container.setAttribute("data-multi-select-container", "true");
      container.setAttribute("data-select-id", selectId);
      container.setAttribute("aria-expanded", "false");

      const toggle = document.createElement("button");
      toggle.type = "button";
      toggle.className = "multi-select-toggle";
      toggle.setAttribute("aria-haspopup", "listbox");
      toggle.setAttribute("aria-controls", `${selectId}_panel`);
      toggle.addEventListener("click", (event) => {
        event.preventDefault();
        toggleOpen(selectEl);
      });
      toggle.addEventListener("keydown", (event) => {
        if (event.key === "ArrowDown") {
          event.preventDefault();
          toggleOpen(selectEl, true);
        } else if (event.key === "Escape") {
          event.preventDefault();
          close(selectEl, { clearSearch: true });
        }
      });

      const valueNode = document.createElement("span");
      valueNode.className = "multi-select-value";
      const arrowNode = document.createElement("span");
      arrowNode.className = "multi-select-arrow";
      arrowNode.setAttribute("aria-hidden", "true");
      arrowNode.textContent = "▾";
      toggle.appendChild(valueNode);
      toggle.appendChild(arrowNode);

      const panel = document.createElement("div");
      panel.className = "multi-select-panel";
      panel.id = `${selectId}_panel`;
      panel.addEventListener("click", (event) => {
        event.stopPropagation();
      });

      container.appendChild(toggle);
      container.appendChild(panel);
      selectEl.insertAdjacentElement("afterend", container);
      selectEl.classList.add("is-multi-hidden");

      state = {
        container,
        toggle,
        panel,
        value: valueNode,
        config: mergedConfig,
        searchTerm: "",
        searchInput: null,
      };
      states.set(selectEl, state);
    } else {
      state.config = mergedConfig;
      selectEl.classList.add("is-multi-hidden");
      state.container.style.display = "";
    }

    render(selectEl);
    syncDisabled(selectEl);
  };

  document.addEventListener("click", (event) => {
    document.querySelectorAll("[data-multi-select-container].is-open").forEach((container) => {
      if (container.contains(event.target)) {
        return;
      }
      const selectId = container.getAttribute("data-select-id");
      if (!selectId) {
        return;
      }
      const selectEl = document.getElementById(selectId);
      if (selectEl) {
        close(selectEl);
      }
    });
  });

  window.MultiSelectFilter = {
    sync,
    destroy,
    close,
    closeAll,
    clear(selectEl) {
      if (!selectEl) {
        return;
      }
      if (selectEl.multiple) {
        Array.from(selectEl.options || []).forEach((option) => {
          option.selected = false;
        });
      } else {
        selectEl.value = "";
      }
      render(selectEl);
    },
    getSelectedValues,
    setSelectedValues,
    syncDisabled,
  };
})();

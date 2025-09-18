(() => {
  const PREVIEW_LIMIT = 5;
  const CATEGORY_TYPE_LABELS = {
    text: 'Texte',
    number: 'Nombre',
    date: 'Date',
    list: 'Liste',
  };
  const numberFormatter = new Intl.NumberFormat('fr-FR');
  const XLSX_ERROR_MESSAGE =
    "L’import de fichiers Excel est indisponible. Vérifiez votre connexion internet.";

  function columnIndexToLetter(index) {
    let result = '';
    let current = index;
    while (current >= 0) {
      result = String.fromCharCode((current % 26) + 65) + result;
      current = Math.floor(current / 26) - 1;
    }
    return result || 'A';
  }

  function formatCategoryType(type) {
    const normalized = typeof type === 'string' ? type.toLowerCase() : 'text';
    return CATEGORY_TYPE_LABELS[normalized] || CATEGORY_TYPE_LABELS.text;
  }

  function getImportApi() {
    if (typeof window !== 'undefined' && window.UManager && window.UManager.importApi) {
      return window.UManager.importApi;
    }
    return null;
  }

  function buildColumnsMetadata(rows) {
    const maxColumns = rows.reduce((max, row) => {
      if (!Array.isArray(row)) {
        return max;
      }
      return Math.max(max, row.length);
    }, 0);

    const columns = [];
    for (let index = 0; index < maxColumns; index += 1) {
      const letter = columnIndexToLetter(index);
      const headerRow = Array.isArray(rows[0]) ? rows[0] : [];
      const headerValue = headerRow[index];
      const header =
        headerValue === undefined || headerValue === null
          ? ''
          : headerValue.toString().trim();

      const samples = [];
      for (let rowIndex = 1; rowIndex < rows.length && samples.length < 3; rowIndex += 1) {
        const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
        const cellValue = row[index];
        const valueString =
          cellValue === undefined || cellValue === null ? '' : cellValue.toString().trim();
        if (valueString && !samples.includes(valueString)) {
          samples.push(valueString);
        }
      }

      columns.push({ index, letter, header, samples });
    }
    return columns;
  }

  function getCategories() {
    const api = getImportApi();
    if (!api || typeof api.getCategories !== 'function') {
      return [];
    }

    try {
      const categories = api.getCategories();
      if (!Array.isArray(categories)) {
        return [];
      }
      return categories
        .map((category) => ({
          id: category && category.id ? category.id : '',
          name: category && category.name ? category.name : '',
          type: category && category.type ? category.type : 'text',
          options: Array.isArray(category && category.options) ? category.options : [],
        }))
        .filter((category) => Boolean(category.id));
    } catch (error) {
      console.warn('Impossible de récupérer les catégories pour l’import :', error);
      return [];
    }
  }

  function getErrorSummaries(errors) {
    if (!Array.isArray(errors)) {
      return [];
    }

    return errors.map((error) => {
      const rowNumber = Number(error && error.row);
      const rowLabel = Number.isFinite(rowNumber) && rowNumber > 0
        ? `Ligne ${numberFormatter.format(rowNumber)}`
        : 'Ligne inconnue';
      const message = error && error.message ? error.message : 'Valeur non reconnue.';
      return `${rowLabel} — ${message}`;
    });
  }

  document.addEventListener('DOMContentLoaded', () => {
    const importPage = document.getElementById('contacts-import');
    if (!importPage) {
      return;
    }

    const fileInput = document.getElementById('contact-import-file');
    const fileSummaryEl = document.getElementById('contact-import-file-summary');
    const optionsSection = document.getElementById('contact-import-options');
    const mappingContainer = document.getElementById('contact-import-mapping');
    const mappingEmptyEl = document.getElementById('contact-import-mapping-empty');
    const hasHeaderCheckbox = document.getElementById('contact-import-has-header');
    const autoMergeCheckbox = document.getElementById('contact-import-auto-merge');
    const feedbackEl = document.getElementById('contact-import-feedback');
    const importButton = document.getElementById('contact-import-submit');
    const resetButton = document.getElementById('contact-import-reset');
    const previewSection = document.getElementById('contact-import-preview');
    const previewTable = previewSection
      ? previewSection.querySelector('.import-preview-table')
      : null;
    const previewHead = previewTable ? previewTable.querySelector('thead') : null;
    const previewBody = previewTable ? previewTable.querySelector('tbody') : null;
    const summaryCountEl = document.getElementById('contact-import-summary-count');

    if (
      !(fileInput instanceof HTMLInputElement) ||
      !(importButton instanceof HTMLButtonElement) ||
      !mappingContainer
    ) {
      return;
    }

    const state = {
      rows: [],
      columns: [],
      mapping: new Map(),
      fileName: '',
      sheetName: '',
      skipHeader: true,
      autoMerge: false,
    };

    function resetFeedback() {
      if (!feedbackEl) {
        return;
      }
      feedbackEl.innerHTML = '';
      feedbackEl.classList.remove(
        'import-feedback--success',
        'import-feedback--error',
        'import-feedback--warning',
      );
    }

    function showFeedback(message, type = 'info', details) {
      if (!feedbackEl) {
        return;
      }
      resetFeedback();
      if (type === 'success') {
        feedbackEl.classList.add('import-feedback--success');
      } else if (type === 'error') {
        feedbackEl.classList.add('import-feedback--error');
      } else if (type === 'warning') {
        feedbackEl.classList.add('import-feedback--warning');
      }

      const text = document.createElement('p');
      text.textContent = message;
      feedbackEl.appendChild(text);

      if (Array.isArray(details) && details.length > 0) {
        const list = document.createElement('ul');
        list.className = 'import-feedback-list';
        details.slice(0, 3).forEach((detail) => {
          const item = document.createElement('li');
          item.textContent = detail;
          list.appendChild(item);
        });
        feedbackEl.appendChild(list);

        if (details.length > 3) {
          const more = document.createElement('p');
          more.className = 'import-feedback-more';
          const remaining = details.length - 3;
          more.textContent = `+ ${numberFormatter.format(remaining)} autre${
            remaining > 1 ? 's' : ''
          } à vérifier.`;
          feedbackEl.appendChild(more);
        }
      }
    }

    function updateOptionsVisibility() {
      if (!optionsSection) {
        return;
      }
      optionsSection.hidden = state.rows.length === 0;
    }

    function updateSummaryCount() {
      if (!summaryCountEl) {
        return;
      }
      if (state.rows.length === 0) {
        summaryCountEl.textContent = '—';
        return;
      }
      const startIndex = state.skipHeader ? 1 : 0;
      const count = Math.max(state.rows.length - startIndex, 0);
      summaryCountEl.textContent = numberFormatter.format(count);
    }

    function clearPreview() {
      if (!previewSection || !previewHead || !previewBody) {
        return;
      }
      previewSection.hidden = true;
      previewHead.innerHTML = '';
      previewBody.innerHTML = '';
    }

    function renderFileSummary() {
      if (!fileSummaryEl) {
        return;
      }

      if (state.rows.length === 0) {
        fileSummaryEl.hidden = true;
        fileSummaryEl.textContent = '';
        return;
      }

      const totalRows = state.rows.length;
      const totalColumns = state.columns.length;
      const sheetInfo = state.sheetName ? ` · Feuille « ${state.sheetName} »` : '';
      fileSummaryEl.hidden = false;
      fileSummaryEl.textContent = `${state.fileName || 'Fichier'} — ${numberFormatter.format(
        totalColumns,
      )} ${totalColumns > 1 ? 'colonnes' : 'colonne'}, ${numberFormatter.format(
        totalRows,
      )} ${totalRows > 1 ? 'lignes' : 'ligne'}${sheetInfo}.`;
    }

    function renderPreview() {
      if (!previewSection || !previewHead || !previewBody) {
        return;
      }

      if (state.rows.length === 0 || state.columns.length === 0) {
        clearPreview();
        return;
      }

      previewHead.innerHTML = '';
      previewBody.innerHTML = '';

      const headerRow = document.createElement('tr');
      state.columns.forEach((column) => {
        const th = document.createElement('th');
        const labelParts = [column.letter];
        if (column.header) {
          labelParts.push(`— ${column.header}`);
        }
        th.textContent = labelParts.join(' ');
        headerRow.appendChild(th);
      });
      previewHead.appendChild(headerRow);

      const startIndex = state.skipHeader ? 1 : 0;
      const sampleRows = state.rows.slice(startIndex, startIndex + PREVIEW_LIMIT);

      if (sampleRows.length === 0) {
        const emptyRow = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = Math.max(state.columns.length, 1);
        cell.textContent = 'Aucune ligne à afficher avec la configuration actuelle.';
        emptyRow.appendChild(cell);
        previewBody.appendChild(emptyRow);
        previewSection.hidden = false;
        return;
      }

      sampleRows.forEach((row) => {
        const tr = document.createElement('tr');
        state.columns.forEach((column) => {
          const td = document.createElement('td');
          const value = Array.isArray(row) ? row[column.index] : undefined;
          td.textContent = value === undefined || value === null ? '' : value.toString();
          tr.appendChild(td);
        });
        previewBody.appendChild(tr);
      });

      previewSection.hidden = false;
    }

    function updateActionsState(categoriesList) {
      const categories = Array.isArray(categoriesList) ? categoriesList : getCategories();
      const hasRows = state.rows.length > 0;
      const hasColumns = state.columns.length > 0;
      const hasCategories = categories.length > 0;
      const hasMapping = hasRows
        ? Array.from(state.mapping.values()).some((value) => Number.isInteger(value) && value >= 0)
        : false;

      importButton.disabled = !(hasRows && hasColumns && hasCategories && hasMapping);
      if (resetButton instanceof HTMLButtonElement) {
        resetButton.disabled = !hasRows;
      }
    }

    function renderCategoryMapping() {
      if (!mappingContainer) {
        return;
      }

      mappingContainer.innerHTML = '';
      if (mappingEmptyEl) {
        mappingEmptyEl.hidden = true;
      }

      if (state.rows.length === 0) {
        updateActionsState([]);
        return;
      }

      if (state.columns.length === 0) {
        if (mappingEmptyEl) {
          mappingEmptyEl.hidden = false;
          mappingEmptyEl.textContent =
            'Aucune colonne exploitable n’a été détectée dans ce fichier Excel.';
        }
        updateActionsState([]);
        return;
      }

      const categories = getCategories();
      const allowedIds = new Set(categories.map((category) => category.id));
      Array.from(state.mapping.keys()).forEach((categoryId) => {
        if (!allowedIds.has(categoryId)) {
          state.mapping.delete(categoryId);
        }
      });

      if (categories.length === 0) {
        if (mappingEmptyEl) {
          mappingEmptyEl.hidden = false;
          mappingEmptyEl.textContent =
            'Ajoutez vos catégories dans l’onglet « Catégories » pour pouvoir importer vos contacts.';
        }
        updateActionsState(categories);
        return;
      }

      const fragment = document.createDocumentFragment();
      categories.forEach((category) => {
        const row = document.createElement('div');
        row.className = 'import-mapping-row';

        const label = document.createElement('div');
        label.className = 'import-mapping-label';

        const title = document.createElement('strong');
        title.textContent = category.name || 'Catégorie';
        label.appendChild(title);

        const meta = document.createElement('span');
        meta.className = 'import-mapping-meta';
        meta.textContent = `Type : ${formatCategoryType(category.type)}`;
        label.appendChild(meta);

        const select = document.createElement('select');
        select.className = 'import-mapping-select';
        select.dataset.categoryId = category.id;

        const noneOption = document.createElement('option');
        noneOption.value = '';
        noneOption.textContent = 'Ne pas importer';
        select.appendChild(noneOption);

        state.columns.forEach((column) => {
          const option = document.createElement('option');
          option.value = column.index.toString();
          const labelParts = [column.letter];
          if (column.header) {
            labelParts.push(`— ${column.header}`);
          }
          if (column.samples.length > 0) {
            labelParts.push(`(ex. ${column.samples.slice(0, 2).join(', ')})`);
          }
          option.textContent = labelParts.join(' ');
          if (state.mapping.get(category.id) === column.index) {
            option.selected = true;
          }
          select.appendChild(option);
        });

        select.addEventListener('change', handleMappingChange);

        row.append(label, select);
        fragment.appendChild(row);
      });

      mappingContainer.appendChild(fragment);
      updateActionsState(categories);
    }

    function handleMappingChange(event) {
      const select = event.currentTarget;
      if (!(select instanceof HTMLSelectElement)) {
        return;
      }
      const categoryId = select.dataset.categoryId || '';
      if (!categoryId) {
        return;
      }

      const value = select.value;
      if (!value) {
        state.mapping.delete(categoryId);
        updateActionsState();
        return;
      }

      const parsed = Number(value);
      if (Number.isInteger(parsed) && parsed >= 0) {
        state.mapping.set(categoryId, parsed);
      } else {
        state.mapping.delete(categoryId);
      }
      updateActionsState();
    }

    function synchronizeView() {
      renderFileSummary();
      updateOptionsVisibility();
      renderCategoryMapping();
      renderPreview();
      updateSummaryCount();
      updateActionsState();
    }

    function handleHeaderToggle() {
      if (!(hasHeaderCheckbox instanceof HTMLInputElement)) {
        return;
      }
      state.skipHeader = hasHeaderCheckbox.checked;
      renderPreview();
      updateSummaryCount();
    }

    function handleAutoMergeToggle() {
      if (!(autoMergeCheckbox instanceof HTMLInputElement)) {
        return;
      }
      state.autoMerge = autoMergeCheckbox.checked;
    }

    function resetState(options = {}) {
      const keepFileInput = Boolean(options.keepFileInput);
      const keepAutoMerge = Boolean(options.keepAutoMerge);
      state.rows = [];
      state.columns = [];
      state.mapping.clear();
      state.fileName = '';
      state.sheetName = '';
      state.skipHeader = true;
      if (!keepAutoMerge) {
        state.autoMerge = false;
      }

      if (!keepFileInput) {
        fileInput.value = '';
      }
      if (hasHeaderCheckbox instanceof HTMLInputElement) {
        hasHeaderCheckbox.checked = true;
      }
      if (autoMergeCheckbox instanceof HTMLInputElement) {
        autoMergeCheckbox.checked = state.autoMerge;
      }
      if (fileSummaryEl) {
        fileSummaryEl.hidden = true;
        fileSummaryEl.textContent = '';
      }
      if (mappingEmptyEl) {
        mappingEmptyEl.hidden = true;
      }
      mappingContainer.innerHTML = '';
      clearPreview();
      updateOptionsVisibility();
      updateSummaryCount();
      updateActionsState();
    }

    async function handleImportClick() {
      const api = getImportApi();
      if (!api || typeof api.importContacts !== 'function') {
        showFeedback("L’importation est indisponible pour le moment.", 'error');
        return;
      }
      if (state.rows.length === 0 || state.columns.length === 0) {
        showFeedback('Veuillez sélectionner un fichier Excel avant de lancer l’import.', 'warning');
        return;
      }
      const mappingEntries = Array.from(state.mapping.entries()).filter(([, value]) =>
        Number.isInteger(value) && value >= 0,
      );
      if (mappingEntries.length === 0) {
        showFeedback('Associez au moins une catégorie à une colonne pour lancer l’import.', 'warning');
        return;
      }

      importButton.disabled = true;
      showFeedback('Import des contacts en cours…', 'info');

      try {
        const payload = {
          rows: state.rows,
          mapping: Object.fromEntries(mappingEntries),
          skipHeader: state.skipHeader,
          fileName: state.fileName,
          autoMerge: state.autoMerge,
        };
        const result = await Promise.resolve(api.importContacts(payload));
        const importedCountRaw =
          result && result.importedCount != null ? Number(result.importedCount) : 0;
        const mergedCountRaw =
          result && result.mergedCount != null ? Number(result.mergedCount) : 0;
        const skippedRaw =
          result && result.skippedEmptyCount != null ? Number(result.skippedEmptyCount) : 0;
        const errorRaw = result && result.errorCount != null ? Number(result.errorCount) : 0;
        const duplicatesRaw =
          result && result.duplicatesDetected != null ? Number(result.duplicatesDetected) : mergedCountRaw;
        const importedCount = Number.isFinite(importedCountRaw) ? importedCountRaw : 0;
        const mergedCount = Number.isFinite(mergedCountRaw) ? mergedCountRaw : 0;
        const skippedEmptyCount = Number.isFinite(skippedRaw) ? skippedRaw : 0;
        const errorCount = Number.isFinite(errorRaw) ? errorRaw : 0;
        const duplicatesDetected = Number.isFinite(duplicatesRaw) ? duplicatesRaw : 0;
        const autoMergeApplied =
          typeof (result && result.autoMergeApplied) === 'boolean'
            ? Boolean(result.autoMergeApplied)
            : Boolean(state.autoMerge);
        const summaryParts = [];
        if (importedCount > 0) {
          summaryParts.push(
            `${numberFormatter.format(importedCount)} contact${importedCount > 1 ? 's' : ''} importé${
              importedCount > 1 ? 's' : ''
            }`,
          );
        }
        if (autoMergeApplied && mergedCount > 0) {
          summaryParts.push(
            `${numberFormatter.format(mergedCount)} contact${mergedCount > 1 ? 's' : ''} fusionné${
              mergedCount > 1 ? 's' : ''
            }`,
          );
        }
        if (!autoMergeApplied && duplicatesDetected > 0) {
          summaryParts.push(
            `${numberFormatter.format(duplicatesDetected)} doublon${
              duplicatesDetected > 1 ? 's' : ''
            } détecté${duplicatesDetected > 1 ? 's' : ''} à fusionner manuellement`,
          );
        }
        if (skippedEmptyCount > 0) {
          summaryParts.push(
            `${numberFormatter.format(skippedEmptyCount)} ligne${
              skippedEmptyCount > 1 ? 's' : ''
            } ignorée${skippedEmptyCount > 1 ? 's' : ''}`,
          );
        }
        if (errorCount > 0) {
          summaryParts.push(
            `${numberFormatter.format(errorCount)} ligne${errorCount > 1 ? 's' : ''} à corriger`,
          );
        }
        if (summaryParts.length === 0) {
          summaryParts.push('Aucun contact ajouté ou fusionné.');
        }

        let feedbackType = 'success';
        const totalIntegrated = importedCount + (autoMergeApplied ? mergedCount : 0);
        if (totalIntegrated === 0) {
          feedbackType = errorCount > 0 ? 'error' : 'warning';
        } else if (errorCount > 0 || skippedEmptyCount > 0) {
          feedbackType = 'warning';
        }
        if (!autoMergeApplied && duplicatesDetected > 0) {
          feedbackType = 'warning';
        }

        const errorSummaries = getErrorSummaries(result && result.errors);
        showFeedback(summaryParts.join(' · '), feedbackType, errorSummaries);
      } catch (error) {
        console.error('Erreur lors de l’import des contacts :', error);
        showFeedback('Une erreur est survenue lors de l’importation des contacts.', 'error');
      } finally {
        updateActionsState();
      }
    }

    function handleReset() {
      resetState();
      showFeedback('Sélectionnez un fichier Excel pour commencer.', 'info');
    }

    function handlePageChanged(event) {
      if (!event || !event.detail) {
        return;
      }
      if (event.detail.pageId === 'contacts-import') {
        synchronizeView();
      }
    }

    async function handleFileChange() {
      if (!fileInput.files || fileInput.files.length === 0) {
        resetState({ keepFileInput: true, keepAutoMerge: true });
        showFeedback('Sélectionnez un fichier Excel pour commencer.', 'info');
        return;
      }

      const file = fileInput.files[0];
      showFeedback(`Analyse de « ${file.name} » en cours…`, 'info');

      try {
        const { rows, sheetName } = await loadWorkbook(file);
        state.rows = Array.isArray(rows) ? rows : [];
        state.columns = buildColumnsMetadata(state.rows);
        state.mapping.clear();
        state.fileName = file.name || '';
        state.sheetName = sheetName || '';
        state.skipHeader = true;
        if (hasHeaderCheckbox instanceof HTMLInputElement) {
          hasHeaderCheckbox.checked = true;
        }

        synchronizeView();

        if (state.rows.length === 0) {
          showFeedback('Le fichier ne contient aucune donnée exploitable.', 'warning');
        } else {
          showFeedback('Fichier analysé. Associez vos catégories aux colonnes détectées.', 'success');
        }
      } catch (error) {
        console.error('Impossible de lire le fichier importé :', error);
        resetState();
        showFeedback(
          'Impossible de lire ce fichier. Vérifiez qu’il s’agit bien d’un fichier Excel valide.',
          'error',
        );
      }
    }

    function loadWorkbook(file) {
      return new Promise((resolve, reject) => {
        if (!window.XLSX || typeof window.XLSX.read !== 'function') {
          reject(new Error('Bibliothèque XLSX indisponible.'));
          return;
        }

        const reader = new FileReader();
        reader.onerror = () => {
          reject(new Error('Lecture du fichier impossible.'));
        };
        reader.onload = () => {
          try {
            const workbook = window.XLSX.read(reader.result, { type: 'array', raw: false });
            const sheetName = workbook.SheetNames[0];
            if (!sheetName) {
              resolve({ rows: [], sheetName: '' });
              return;
            }
            const sheet = workbook.Sheets[sheetName];
            const rows = window.XLSX.utils.sheet_to_json(sheet, {
              header: 1,
              raw: false,
              blankrows: true,
              defval: '',
            });
            resolve({ rows, sheetName });
          } catch (error) {
            reject(error);
          }
        };
        reader.readAsArrayBuffer(file);
      });
    }

    if (!window.XLSX || typeof window.XLSX.read !== 'function') {
      fileInput.disabled = true;
      importButton.disabled = true;
      if (resetButton instanceof HTMLButtonElement) {
        resetButton.disabled = true;
      }
      showFeedback(XLSX_ERROR_MESSAGE, 'error');
      return;
    }

    fileInput.addEventListener('change', handleFileChange);
    if (hasHeaderCheckbox instanceof HTMLInputElement) {
      hasHeaderCheckbox.addEventListener('change', handleHeaderToggle);
    }
    if (autoMergeCheckbox instanceof HTMLInputElement) {
      autoMergeCheckbox.checked = state.autoMerge;
      autoMergeCheckbox.addEventListener('change', handleAutoMergeToggle);
    }
    importButton.addEventListener('click', handleImportClick);
    if (resetButton instanceof HTMLButtonElement) {
      resetButton.addEventListener('click', handleReset);
    }

    document.addEventListener('umanager:page-changed', handlePageChanged);

    synchronizeView();
    if (state.rows.length === 0) {
      showFeedback('Sélectionnez un fichier Excel pour commencer.', 'info');
    }
  });
})();

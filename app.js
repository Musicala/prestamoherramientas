'use strict';

/* ============================================================================
   PRÉSTAMOS MUSICALA · app.js
   - Registro de préstamos
   - Historial y activos
   - Devoluciones con notas
   - Catálogos de personas, categorías y elementos
   - localStorage
   - Exportar / importar
   - Filtros y búsqueda
============================================================================ */

/* =========================
   STORAGE KEYS
========================= */
const STORAGE_KEYS = {
  loans: 'musicala_prestamos_v2',
  catalogs: 'musicala_prestamos_catalogos_v2'
};

/* =========================
   HELPERS DOM
========================= */
const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));

/* =========================
   ESTADO
========================= */
const state = {
  loans: [],
  currentTab: 'active',
  filters: {
    search: '',
    type: 'Todos',
    category: 'Todos',
    status: 'Todos'
  },
  catalogs: {
    people: [],
    categories: [],
    items: []
  }
};

/* =========================
   ELEMENTOS DOM
========================= */
const loanForm = $('#loanForm');
const btnClearForm = $('#btnClearForm');

const personType = $('#personType');
const personName = $('#personName');
const itemCategory = $('#itemCategory');
const itemName = $('#itemName');
const quantity = $('#quantity');
const loanNotes = $('#loanNotes');

const btnQuickAddPerson = $('#btnQuickAddPerson');
const btnQuickAddCategory = $('#btnQuickAddCategory');
const btnQuickAddItem = $('#btnQuickAddItem');

const activeList = $('#activeList');
const historyTableBody = $('#historyTableBody');

const statActive = $('#statActive');
const statReturned = $('#statReturned');
const statIssues = $('#statIssues');
const statCatalogPeople = $('#statCatalogPeople');
const statCatalogCategories = $('#statCatalogCategories');
const statCatalogItems = $('#statCatalogItems');

const searchInput = $('#searchInput');
const filterType = $('#filterType');
const filterCategory = $('#filterCategory');
const filterStatus = $('#filterStatus');

const tabBtnActive = $('#tabBtnActive');
const tabBtnHistory = $('#tabBtnHistory');
const panelActive = $('#panelActive');
const panelHistory = $('#panelHistory');

const returnModal = $('#returnModal');
const returnForm = $('#returnForm');
const returnLoanId = $('#returnLoanId');
const returnCondition = $('#returnCondition');
const returnNotes = $('#returnNotes');
const btnCloseModal = $('#btnCloseModal');
const btnCancelModal = $('#btnCancelModal');

const catalogsModal = $('#catalogsModal');
const btnOpenCatalogs = $('#btnOpenCatalogs');
const btnCloseCatalogsModal = $('#btnCloseCatalogsModal');

const personCatalogForm = $('#personCatalogForm');
const categoryCatalogForm = $('#categoryCatalogForm');
const itemCatalogForm = $('#itemCatalogForm');

const catalogPersonType = $('#catalogPersonType');
const catalogPersonName = $('#catalogPersonName');
const catalogCategoryName = $('#catalogCategoryName');
const catalogItemCategory = $('#catalogItemCategory');
const catalogItemName = $('#catalogItemName');

const peopleCatalogList = $('#peopleCatalogList');
const categoriesCatalogList = $('#categoriesCatalogList');
const itemsCatalogList = $('#itemsCatalogList');

const btnExportJson = $('#btnExportJson');
const btnExportCsv = $('#btnExportCsv');
const importFile = $('#importFile');

const activeLoanCardTemplate = $('#activeLoanCardTemplate');
const catalogChipTemplate = $('#catalogChipTemplate');

/* =========================
   INIT
========================= */
document.addEventListener('DOMContentLoaded', init);

function init() {
  loadCatalogs();
  loadLoans();
  seedDefaultsIfNeeded();
  bindEvents();
  renderAll();
}

/* =========================
   EVENTOS
========================= */
function bindEvents() {
  loanForm?.addEventListener('submit', handleCreateLoan);
  btnClearForm?.addEventListener('click', handleClearForm);

  tabBtnActive?.addEventListener('click', () => switchTab('active'));
  tabBtnHistory?.addEventListener('click', () => switchTab('history'));

  searchInput?.addEventListener('input', (e) => {
    state.filters.search = e.target.value.trim().toLowerCase();
    renderAll();
  });

  filterType?.addEventListener('change', (e) => {
    state.filters.type = e.target.value;
    renderAll();
  });

  filterCategory?.addEventListener('change', (e) => {
    state.filters.category = e.target.value;
    renderAll();
  });

  filterStatus?.addEventListener('change', (e) => {
    state.filters.status = e.target.value;
    renderAll();
  });

  itemCategory?.addEventListener('change', () => {
    populateItemSelect(itemCategory.value, '');
  });

  personType?.addEventListener('change', () => {
    populatePersonSelect(personType.value, '');
  });

  returnForm?.addEventListener('submit', handleSubmitReturn);
  btnCloseModal?.addEventListener('click', closeReturnModal);
  btnCancelModal?.addEventListener('click', closeReturnModal);

  btnOpenCatalogs?.addEventListener('click', openCatalogsModal);
  btnCloseCatalogsModal?.addEventListener('click', closeCatalogsModal);

  personCatalogForm?.addEventListener('submit', handleAddPersonToCatalog);
  categoryCatalogForm?.addEventListener('submit', handleAddCategoryToCatalog);
  itemCatalogForm?.addEventListener('submit', handleAddItemToCatalog);

  btnQuickAddPerson?.addEventListener('click', handleQuickAddPerson);
  btnQuickAddCategory?.addEventListener('click', handleQuickAddCategory);
  btnQuickAddItem?.addEventListener('click', handleQuickAddItem);

  btnExportJson?.addEventListener('click', exportJSON);
  btnExportCsv?.addEventListener('click', exportCSV);
  importFile?.addEventListener('change', handleImportJSON);

  document.addEventListener('click', handleDelegatedClicks);

  bindDialogBackdropClose(returnModal, closeReturnModal);
  bindDialogBackdropClose(catalogsModal, closeCatalogsModal);
}

function bindDialogBackdropClose(dialog, onClose) {
  if (!dialog) return;

  dialog.addEventListener('click', (e) => {
    const rect = dialog.getBoundingClientRect();
    const clickedInside =
      e.clientX >= rect.left &&
      e.clientX <= rect.right &&
      e.clientY >= rect.top &&
      e.clientY <= rect.bottom;

    if (!clickedInside) onClose();
  });
}

function handleDelegatedClicks(e) {
  const returnBtn = e.target.closest('.btn-return');
  if (returnBtn) {
    openReturnModal(returnBtn.dataset.id);
    return;
  }

  const deleteBtn = e.target.closest('.btn-delete');
  if (deleteBtn) {
    deleteLoan(deleteBtn.dataset.id);
    return;
  }

  const reopenBtn = e.target.closest('.btn-reopen');
  if (reopenBtn) {
    reopenLoan(reopenBtn.dataset.id);
    return;
  }

  const deleteCatalogBtn = e.target.closest('.btn-delete-catalog-item');
  if (deleteCatalogBtn) {
    const type = deleteCatalogBtn.dataset.catalogType;
    const id = deleteCatalogBtn.dataset.catalogId;
    deleteCatalogEntry(type, id);
  }
}

/* =========================
   CRUD PRESTAMOS
========================= */
function handleCreateLoan(e) {
  e.preventDefault();

  const formData = new FormData(loanForm);

  const loanPersonType = String(formData.get('personType') || '').trim();
  const loanPersonName = String(formData.get('personName') || '').trim();
  const loanItemCategory = String(formData.get('itemCategory') || '').trim();
  const loanItemName = String(formData.get('itemName') || '').trim();
  const loanQuantity = Number(formData.get('quantity') || 1);
  const notes = String(formData.get('loanNotes') || '').trim();

  if (!loanPersonType || !loanPersonName || !loanItemCategory || !loanItemName || loanQuantity < 1) {
    alert('Completa los campos obligatorios del préstamo.');
    return;
  }

  const personExists = state.catalogs.people.some(
    p => p.name === loanPersonName && p.type === loanPersonType
  );

  const categoryExists = state.catalogs.categories.some(
    c => c.name === loanItemCategory
  );

  const itemExists = state.catalogs.items.some(
    i => i.name === loanItemName && i.category === loanItemCategory
  );

  if (!personExists || !categoryExists || !itemExists) {
    alert('La persona, categoría o elemento ya no existe en las listas guardadas. Revisa el catálogo.');
    return;
  }

  const now = new Date();

  const newLoan = {
    id: generateLoanId(now),
    personType: loanPersonType,
    personName: loanPersonName,
    itemCategory: loanItemCategory,
    itemName: loanItemName,
    quantity: loanQuantity,
    loanDateISO: now.toISOString(),
    loanDateLabel: formatDateTime(now),
    status: 'Prestado',
    loanNotes: notes,
    returnDateISO: '',
    returnDateLabel: '',
    returnCondition: '',
    returnNotes: '',
    createdAt: now.toISOString(),
    updatedAt: now.toISOString()
  };

  state.loans.unshift(newLoan);
  persistLoans();
  renderAll();
  resetLoanForm();
}

function handleSubmitReturn(e) {
  e.preventDefault();

  const loanId = returnLoanId?.value;
  if (!loanId) return;

  const loan = state.loans.find(item => item.id === loanId);
  if (!loan) {
    closeReturnModal();
    return;
  }

  const condition = returnCondition?.value || 'Devuelto';
  const notes = (returnNotes?.value || '').trim();
  const now = new Date();

  loan.status = condition;
  loan.returnCondition = condition;
  loan.returnNotes = notes;
  loan.returnDateISO = now.toISOString();
  loan.returnDateLabel = formatDateTime(now);
  loan.updatedAt = now.toISOString();

  persistLoans();
  renderAll();
  closeReturnModal();
}

function deleteLoan(loanId) {
  const loan = state.loans.find(item => item.id === loanId);
  if (!loan) return;

  const confirmed = window.confirm(
    `¿Eliminar el registro de ${loan.personName} (${loan.itemName})?`
  );

  if (!confirmed) return;

  state.loans = state.loans.filter(item => item.id !== loanId);
  persistLoans();
  renderAll();
}

function reopenLoan(loanId) {
  const loan = state.loans.find(item => item.id === loanId);
  if (!loan) return;

  loan.status = 'Prestado';
  loan.returnCondition = '';
  loan.returnNotes = '';
  loan.returnDateISO = '';
  loan.returnDateLabel = '';
  loan.updatedAt = new Date().toISOString();

  persistLoans();
  renderAll();
}

/* =========================
   CRUD CATALOGOS
========================= */
function handleAddPersonToCatalog(e) {
  e.preventDefault();

  const type = String(catalogPersonType?.value || '').trim();
  const name = String(catalogPersonName?.value || '').trim();

  if (!type || !name) {
    alert('Completa tipo y nombre de la persona.');
    return;
  }

  const exists = state.catalogs.people.some(
    item => item.type === type && normalizeText(item.name) === normalizeText(name)
  );

  if (exists) {
    alert('Esa persona ya está guardada en la lista.');
    return;
  }

  state.catalogs.people.unshift({
    id: generateCatalogId('PER'),
    type,
    name
  });

  persistCatalogs();
  renderCatalogs();
  syncFormSelects();
  syncFilters();
  personCatalogForm.reset();
  if (personType?.value === type) {
    populatePersonSelect(type, name);
  }
}

function handleAddCategoryToCatalog(e) {
  e.preventDefault();

  const name = String(catalogCategoryName?.value || '').trim();

  if (!name) {
    alert('Escribe el nombre de la categoría.');
    return;
  }

  const exists = state.catalogs.categories.some(
    item => normalizeText(item.name) === normalizeText(name)
  );

  if (exists) {
    alert('Esa categoría ya existe.');
    return;
  }

  state.catalogs.categories.unshift({
    id: generateCatalogId('CAT'),
    name
  });

  persistCatalogs();
  renderCatalogs();
  syncFormSelects();
  syncFilters();
  categoryCatalogForm.reset();
  populateCategorySelect(name);
  populateCatalogItemCategorySelect(name);
}

function handleAddItemToCatalog(e) {
  e.preventDefault();

  const category = String(catalogItemCategory?.value || '').trim();
  const name = String(catalogItemName?.value || '').trim();

  if (!category || !name) {
    alert('Completa categoría y nombre del elemento.');
    return;
  }

  const exists = state.catalogs.items.some(
    item =>
      item.category === category &&
      normalizeText(item.name) === normalizeText(name)
  );

  if (exists) {
    alert('Ese elemento ya existe dentro de esa categoría.');
    return;
  }

  state.catalogs.items.unshift({
    id: generateCatalogId('ITE'),
    category,
    name
  });

  persistCatalogs();
  renderCatalogs();
  syncFormSelects();
  syncFilters();
  itemCatalogForm.reset();
  populateCatalogItemCategorySelect(category);
  populateCategorySelect(category);
  populateItemSelect(category, name);
}

function deleteCatalogEntry(type, id) {
  if (!type || !id) return;

  if (type === 'people') {
    const target = state.catalogs.people.find(item => item.id === id);
    if (!target) return;

    const linkedLoans = state.loans.some(
      loan => loan.personName === target.name && loan.personType === target.type
    );

    if (linkedLoans) {
      alert('No puedes eliminar esa persona porque ya aparece en préstamos registrados.');
      return;
    }

    const confirmed = window.confirm(`¿Eliminar a ${target.name} de la lista?`);
    if (!confirmed) return;

    state.catalogs.people = state.catalogs.people.filter(item => item.id !== id);
  }

  if (type === 'categories') {
    const target = state.catalogs.categories.find(item => item.id === id);
    if (!target) return;

    const hasItems = state.catalogs.items.some(item => item.category === target.name);
    const linkedLoans = state.loans.some(loan => loan.itemCategory === target.name);

    if (hasItems || linkedLoans) {
      alert('No puedes eliminar esa categoría porque tiene elementos o ya aparece en préstamos.');
      return;
    }

    const confirmed = window.confirm(`¿Eliminar la categoría ${target.name}?`);
    if (!confirmed) return;

    state.catalogs.categories = state.catalogs.categories.filter(item => item.id !== id);
  }

  if (type === 'items') {
    const target = state.catalogs.items.find(item => item.id === id);
    if (!target) return;

    const linkedLoans = state.loans.some(
      loan => loan.itemName === target.name && loan.itemCategory === target.category
    );

    if (linkedLoans) {
      alert('No puedes eliminar ese elemento porque ya aparece en préstamos registrados.');
      return;
    }

    const confirmed = window.confirm(`¿Eliminar el elemento ${target.name}?`);
    if (!confirmed) return;

    state.catalogs.items = state.catalogs.items.filter(item => item.id !== id);
  }

  persistCatalogs();
  renderCatalogs();
  syncFormSelects();
  syncFilters();
  renderAll();
}

/* =========================
   QUICK ADD
========================= */
function handleQuickAddPerson() {
  openCatalogsModal();
  const currentType = String(personType?.value || '').trim();
  if (catalogPersonType && currentType) catalogPersonType.value = currentType;
  catalogPersonName?.focus();
}

function handleQuickAddCategory() {
  openCatalogsModal();
  catalogCategoryName?.focus();
}

function handleQuickAddItem() {
  openCatalogsModal();
  const currentCategory = String(itemCategory?.value || '').trim();
  if (catalogItemCategory && currentCategory) catalogItemCategory.value = currentCategory;
  catalogItemName?.focus();
}

/* =========================
   MODALES
========================= */
function openReturnModal(loanId) {
  const loan = state.loans.find(item => item.id === loanId);
  if (!loan) return;

  if (returnLoanId) returnLoanId.value = loan.id;
  if (returnCondition) returnCondition.value = loan.status === 'Con novedad' ? 'Con novedad' : 'Devuelto';
  if (returnNotes) returnNotes.value = loan.returnNotes || '';

  if (typeof returnModal?.showModal === 'function' && !returnModal.open) {
    returnModal.showModal();
  }
}

function closeReturnModal() {
  if (returnForm) returnForm.reset();
  if (returnLoanId) returnLoanId.value = '';

  if (returnModal?.open) {
    returnModal.close();
  }
}

function openCatalogsModal() {
  if (typeof catalogsModal?.showModal === 'function' && !catalogsModal.open) {
    catalogsModal.showModal();
  }
}

function closeCatalogsModal() {
  if (catalogsModal?.open) {
    catalogsModal.close();
  }
}

/* =========================
   STORAGE
========================= */
function loadLoans() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.loans);
    if (!raw) {
      const legacyRaw = localStorage.getItem('musicala_prestamos_v1');
      if (legacyRaw) {
        const legacyParsed = JSON.parse(legacyRaw);
        state.loans = Array.isArray(legacyParsed)
          ? legacyParsed.map(normalizeImportedLoan).filter(Boolean)
          : [];
        persistLoans();
      } else {
        state.loans = [];
      }
      return;
    }

    const parsed = JSON.parse(raw);
    state.loans = Array.isArray(parsed)
      ? parsed.map(normalizeImportedLoan).filter(Boolean)
      : [];
  } catch (error) {
    console.error('Error cargando préstamos desde localStorage:', error);
    state.loans = [];
  }
}

function loadCatalogs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.catalogs);
    if (!raw) {
      state.catalogs = {
        people: [],
        categories: [],
        items: []
      };
      return;
    }

    const parsed = JSON.parse(raw);
    state.catalogs = normalizeCatalogs(parsed);
  } catch (error) {
    console.error('Error cargando catálogos:', error);
    state.catalogs = {
      people: [],
      categories: [],
      items: []
    };
  }
}

function persistLoans() {
  try {
    localStorage.setItem(STORAGE_KEYS.loans, JSON.stringify(state.loans));
  } catch (error) {
    console.error('Error guardando préstamos en localStorage:', error);
    alert('No se pudo guardar en localStorage. Revisa el espacio del navegador.');
  }
}

function persistCatalogs() {
  try {
    localStorage.setItem(STORAGE_KEYS.catalogs, JSON.stringify(state.catalogs));
  } catch (error) {
    console.error('Error guardando catálogos en localStorage:', error);
    alert('No se pudieron guardar las listas en localStorage.');
  }
}

/* =========================
   DATA SEED
========================= */
function seedDefaultsIfNeeded() {
  let changed = false;

  if (!state.catalogs.categories.length) {
    state.catalogs.categories = [
      { id: generateCatalogId('CAT'), name: 'Audio' },
      { id: generateCatalogId('CAT'), name: 'Cables' },
      { id: generateCatalogId('CAT'), name: 'Instrumentos' },
      { id: generateCatalogId('CAT'), name: 'Accesorios' }
    ];
    changed = true;
  }

  if (!state.catalogs.items.length) {
    state.catalogs.items = [
      { id: generateCatalogId('ITE'), category: 'Audio', name: 'Micrófono inalámbrico' },
      { id: generateCatalogId('ITE'), category: 'Audio', name: 'Parlante' },
      { id: generateCatalogId('ITE'), category: 'Cables', name: 'Cable XLR' },
      { id: generateCatalogId('ITE'), category: 'Cables', name: 'Extensión' },
      { id: generateCatalogId('ITE'), category: 'Instrumentos', name: 'Guitarra' },
      { id: generateCatalogId('ITE'), category: 'Instrumentos', name: 'Ukelele' },
      { id: generateCatalogId('ITE'), category: 'Accesorios', name: 'Atril' }
    ];
    changed = true;
  }

  if (changed) {
    persistCatalogs();
  }
}

/* =========================
   RENDER GENERAL
========================= */
function renderAll() {
  renderStats();
  renderCatalogStats();
  renderCatalogs();
  syncFormSelects();
  syncFilters();
  renderActiveLoans();
  renderHistoryTable();
}

/* =========================
   RENDER STATS
========================= */
function renderStats() {
  const activeCount = state.loans.filter(item => item.status === 'Prestado').length;
  const returnedCount = state.loans.filter(item => item.status === 'Devuelto').length;
  const issuesCount = state.loans.filter(item => item.status === 'Con novedad').length;

  if (statActive) statActive.textContent = String(activeCount);
  if (statReturned) statReturned.textContent = String(returnedCount);
  if (statIssues) statIssues.textContent = String(issuesCount);
}

function renderCatalogStats() {
  if (statCatalogPeople) statCatalogPeople.textContent = String(state.catalogs.people.length);
  if (statCatalogCategories) statCatalogCategories.textContent = String(state.catalogs.categories.length);
  if (statCatalogItems) statCatalogItems.textContent = String(state.catalogs.items.length);
}

/* =========================
   RENDER CATALOGOS
========================= */
function renderCatalogs() {
  renderPeopleCatalog();
  renderCategoriesCatalog();
  renderItemsCatalog();
}

function renderPeopleCatalog() {
  if (!peopleCatalogList) return;

  const items = [...state.catalogs.people].sort(comparePeople);

  if (!items.length) {
    peopleCatalogList.innerHTML = `
      <article class="empty-state small-empty">
        <p>No hay personas guardadas.</p>
      </article>
    `;
    return;
  }

  peopleCatalogList.innerHTML = '';
  const fragment = document.createDocumentFragment();

  items.forEach(item => {
    fragment.appendChild(
      buildCatalogChip({
        type: 'people',
        id: item.id,
        title: item.name,
        meta: item.type
      })
    );
  });

  peopleCatalogList.appendChild(fragment);
}

function renderCategoriesCatalog() {
  if (!categoriesCatalogList) return;

  const items = [...state.catalogs.categories].sort((a, b) =>
    a.name.localeCompare(b.name, 'es', { sensitivity: 'base' })
  );

  if (!items.length) {
    categoriesCatalogList.innerHTML = `
      <article class="empty-state small-empty">
        <p>No hay categorías guardadas.</p>
      </article>
    `;
    return;
  }

  categoriesCatalogList.innerHTML = '';
  const fragment = document.createDocumentFragment();

  items.forEach(item => {
    const count = state.catalogs.items.filter(catItem => catItem.category === item.name).length;
    fragment.appendChild(
      buildCatalogChip({
        type: 'categories',
        id: item.id,
        title: item.name,
        meta: `${count} elemento${count === 1 ? '' : 's'}`
      })
    );
  });

  categoriesCatalogList.appendChild(fragment);
}

function renderItemsCatalog() {
  if (!itemsCatalogList) return;

  const items = [...state.catalogs.items].sort(compareItems);

  if (!items.length) {
    itemsCatalogList.innerHTML = `
      <article class="empty-state small-empty">
        <p>No hay elementos guardados.</p>
      </article>
    `;
    return;
  }

  itemsCatalogList.innerHTML = '';
  const fragment = document.createDocumentFragment();

  items.forEach(item => {
    fragment.appendChild(
      buildCatalogChip({
        type: 'items',
        id: item.id,
        title: item.name,
        meta: item.category
      })
    );
  });

  itemsCatalogList.appendChild(fragment);
}

function buildCatalogChip({ type, id, title, meta }) {
  if (!catalogChipTemplate) {
    const fallback = document.createElement('article');
    fallback.className = 'catalog-chip';
    fallback.innerHTML = `
      <div class="catalog-chip-copy">
        <strong class="catalog-chip-title">${escapeHTML(title)}</strong>
        <small class="catalog-chip-meta">${escapeHTML(meta || '')}</small>
      </div>
      <button
        type="button"
        class="icon-btn icon-btn-danger btn-delete-catalog-item"
        data-catalog-type="${escapeHTML(type)}"
        data-catalog-id="${escapeHTML(id)}"
        aria-label="Eliminar"
      >×</button>
    `;
    return fallback;
  }

  const clone = catalogChipTemplate.content.cloneNode(true);
  $('.catalog-chip-title', clone).textContent = title;
  $('.catalog-chip-meta', clone).textContent = meta || '';

  const btn = $('.btn-delete-catalog-item', clone);
  if (btn) {
    btn.dataset.catalogType = type;
    btn.dataset.catalogId = id;
  }

  const wrapper = document.createElement('div');
  wrapper.appendChild(clone);
  return wrapper.firstElementChild;
}

/* =========================
   RENDER SELECTS / FILTROS
========================= */
function syncFormSelects() {
  const selectedType = personType?.value || '';
  const selectedPerson = personName?.value || '';
  const selectedCategory = itemCategory?.value || '';
  const selectedItem = itemName?.value || '';

  populatePersonSelect(selectedType, selectedPerson);
  populateCategorySelect(selectedCategory);
  populateItemSelect(selectedCategory, selectedItem);
  populateCatalogItemCategorySelect(catalogItemCategory?.value || '');
}

function syncFilters() {
  populateFilterCategories(filterCategory?.value || state.filters.category || 'Todos');
}

function populatePersonSelect(selectedType = '', selectedValue = '') {
  if (!personName) return;

  const availablePeople = selectedType
    ? state.catalogs.people
        .filter(item => item.type === selectedType)
        .sort(comparePeople)
    : [...state.catalogs.people].sort(comparePeople);

  personName.innerHTML = '<option value="">Selecciona una persona</option>';

  availablePeople.forEach(item => {
    const option = document.createElement('option');
    option.value = item.name;
    option.textContent = selectedType ? item.name : `${item.name} · ${item.type}`;
    if (item.name === selectedValue) option.selected = true;
    personName.appendChild(option);
  });
}

function populateCategorySelect(selectedValue = '') {
  if (!itemCategory) return;

  const categories = [...state.catalogs.categories].sort((a, b) =>
    a.name.localeCompare(b.name, 'es', { sensitivity: 'base' })
  );

  itemCategory.innerHTML = '<option value="">Selecciona una categoría</option>';

  categories.forEach(item => {
    const option = document.createElement('option');
    option.value = item.name;
    option.textContent = item.name;
    if (item.name === selectedValue) option.selected = true;
    itemCategory.appendChild(option);
  });
}

function populateItemSelect(category = '', selectedValue = '') {
  if (!itemName) return;

  const items = category
    ? state.catalogs.items
        .filter(item => item.category === category)
        .sort(compareItems)
    : [];

  itemName.innerHTML = '<option value="">Selecciona un elemento</option>';

  items.forEach(item => {
    const option = document.createElement('option');
    option.value = item.name;
    option.textContent = item.name;
    if (item.name === selectedValue) option.selected = true;
    itemName.appendChild(option);
  });
}

function populateCatalogItemCategorySelect(selectedValue = '') {
  if (!catalogItemCategory) return;

  const categories = [...state.catalogs.categories].sort((a, b) =>
    a.name.localeCompare(b.name, 'es', { sensitivity: 'base' })
  );

  catalogItemCategory.innerHTML = '<option value="">Selecciona categoría</option>';

  categories.forEach(item => {
    const option = document.createElement('option');
    option.value = item.name;
    option.textContent = item.name;
    if (item.name === selectedValue) option.selected = true;
    catalogItemCategory.appendChild(option);
  });
}

function populateFilterCategories(selectedValue = 'Todos') {
  if (!filterCategory) return;

  const categories = [...state.catalogs.categories].sort((a, b) =>
    a.name.localeCompare(b.name, 'es', { sensitivity: 'base' })
  );

  filterCategory.innerHTML = '<option value="Todos">Todas las categorías</option>';

  categories.forEach(item => {
    const option = document.createElement('option');
    option.value = item.name;
    option.textContent = item.name;
    if (item.name === selectedValue) option.selected = true;
    filterCategory.appendChild(option);
  });
}

/* =========================
   RENDER PRESTAMOS
========================= */
function renderActiveLoans() {
  if (!activeList) return;

  const filtered = getFilteredLoans().filter(item => item.status === 'Prestado');

  if (!filtered.length) {
    activeList.innerHTML = `
      <article class="empty-state">
        <h3>No hay préstamos activos</h3>
        <p>Cuando registres un préstamo aparecerá aquí para poder marcar su devolución.</p>
      </article>
    `;
    return;
  }

  activeList.innerHTML = '';
  const fragment = document.createDocumentFragment();

  filtered.forEach(loan => {
    fragment.appendChild(buildActiveLoanCard(loan));
  });

  activeList.appendChild(fragment);
}

function buildActiveLoanCard(loan) {
  const statusClass = getStatusClass(loan.status);

  if (!activeLoanCardTemplate) {
    const fallback = document.createElement('article');
    fallback.className = 'loan-card';
    fallback.innerHTML = `
      <div class="loan-card-top">
        <div>
          <h3 class="loan-person">${escapeHTML(loan.personName)}</h3>
          <p class="loan-meta">${escapeHTML(loan.personType)} · ${escapeHTML(loan.id)}</p>
        </div>
        <span class="pill loan-status ${statusClass}">${escapeHTML(loan.status)}</span>
      </div>
      <div class="loan-card-body">
        <p><strong>Categoría:</strong> ${escapeHTML(loan.itemCategory || 'Sin categoría')}</p>
        <p><strong>Elemento:</strong> ${escapeHTML(loan.itemName)}</p>
        <p><strong>Cantidad:</strong> ${escapeHTML(String(loan.quantity))}</p>
        <p><strong>Fecha:</strong> ${escapeHTML(loan.loanDateLabel)}</p>
        <p><strong>Notas:</strong> ${escapeHTML(loan.loanNotes || 'Sin notas')}</p>
      </div>
      <div class="loan-card-actions">
        <button type="button" class="btn btn-primary btn-return" data-id="${escapeHTML(loan.id)}">
          Marcar devolución
        </button>
      </div>
    `;
    return fallback;
  }

  const clone = activeLoanCardTemplate.content.cloneNode(true);

  $('.loan-person', clone).textContent = loan.personName;
  $('.loan-meta', clone).textContent = `${loan.personType} · ${loan.id}`;

  const statusEl = $('.loan-status', clone);
  statusEl.textContent = loan.status;
  statusEl.classList.add(statusClass);

  $('.loan-category', clone).textContent = loan.itemCategory || 'Sin categoría';
  $('.loan-item', clone).textContent = loan.itemName;
  $('.loan-qty', clone).textContent = String(loan.quantity);
  $('.loan-date', clone).textContent = loan.loanDateLabel;
  $('.loan-notes', clone).textContent = loan.loanNotes || 'Sin notas';

  const btnReturn = $('.btn-return', clone);
  if (btnReturn) btnReturn.dataset.id = loan.id;

  const wrapper = document.createElement('div');
  wrapper.appendChild(clone);
  return wrapper.firstElementChild;
}

function renderHistoryTable() {
  if (!historyTableBody) return;

  const filtered = getFilteredLoans();

  if (!filtered.length) {
    historyTableBody.innerHTML = `
      <tr>
        <td colspan="11">
          <div class="empty-state table-empty">
            <h3>Sin resultados</h3>
            <p>No hay registros que coincidan con los filtros actuales.</p>
          </div>
        </td>
      </tr>
    `;
    return;
  }

  historyTableBody.innerHTML = filtered
    .map(loan => {
      const notes = buildNotesSummary(loan);
      const returnInfo = loan.returnDateLabel || 'Pendiente';
      const actionsHTML = loan.status === 'Prestado'
        ? `<button type="button" class="btn btn-primary btn-return" data-id="${escapeHTML(loan.id)}">Devolver</button>`
        : `<button type="button" class="btn btn-secondary btn-reopen" data-id="${escapeHTML(loan.id)}">Reabrir</button>`;

      return `
        <tr>
          <td>${escapeHTML(loan.id)}</td>
          <td>${escapeHTML(loan.personName)}</td>
          <td>${escapeHTML(loan.personType)}</td>
          <td>${escapeHTML(loan.itemCategory || 'Sin categoría')}</td>
          <td>${escapeHTML(loan.itemName)}</td>
          <td>${escapeHTML(String(loan.quantity))}</td>
          <td>${escapeHTML(loan.loanDateLabel)}</td>
          <td>${escapeHTML(returnInfo)}</td>
          <td>${renderStatusBadgeText(loan.status)}</td>
          <td>${escapeHTML(notes)}</td>
          <td>
            <div style="display:flex; gap:8px; flex-wrap:wrap;">
              ${actionsHTML}
              <button type="button" class="btn btn-ghost btn-delete" data-id="${escapeHTML(loan.id)}">Eliminar</button>
            </div>
          </td>
        </tr>
      `;
    })
    .join('');
}

/* =========================
   FILTROS
========================= */
function getFilteredLoans() {
  const search = state.filters.search;
  const type = state.filters.type;
  const category = state.filters.category;
  const status = state.filters.status;

  return [...state.loans]
    .filter(item => {
      if (type !== 'Todos' && item.personType !== type) return false;
      if (category !== 'Todos' && item.itemCategory !== category) return false;
      if (status !== 'Todos' && item.status !== status) return false;

      if (!search) return true;

      const haystack = [
        item.id,
        item.personType,
        item.personName,
        item.itemCategory,
        item.itemName,
        item.loanNotes,
        item.returnNotes,
        item.status,
        item.returnCondition
      ]
        .join(' ')
        .toLowerCase();

      return haystack.includes(search);
    })
    .sort((a, b) => {
      const aTime = new Date(a.createdAt || a.loanDateISO).getTime();
      const bTime = new Date(b.createdAt || b.loanDateISO).getTime();
      return bTime - aTime;
    });
}

/* =========================
   TABS
========================= */
function switchTab(tabName) {
  state.currentTab = tabName;

  const isActiveTab = tabName === 'active';

  tabBtnActive?.classList.toggle('is-active', isActiveTab);
  tabBtnHistory?.classList.toggle('is-active', !isActiveTab);

  tabBtnActive?.setAttribute('aria-selected', String(isActiveTab));
  tabBtnHistory?.setAttribute('aria-selected', String(!isActiveTab));

  panelActive?.classList.toggle('is-active', isActiveTab);
  panelHistory?.classList.toggle('is-active', !isActiveTab);

  if (panelActive) panelActive.hidden = !isActiveTab;
  if (panelHistory) panelHistory.hidden = isActiveTab;
}

/* =========================
   EXPORT / IMPORT
========================= */
function exportJSON() {
  const payload = {
    version: 2,
    exportedAt: new Date().toISOString(),
    catalogs: state.catalogs,
    loans: state.loans
  };

  const fileName = `prestamos-musicala-${getTodayFileSafe()}.json`;
  const content = JSON.stringify(payload, null, 2);
  downloadFile(content, fileName, 'application/json;charset=utf-8;');
}

function exportCSV() {
  const rows = [
    [
      'ID',
      'TipoPersona',
      'Nombre',
      'Categoria',
      'Elemento',
      'Cantidad',
      'FechaPrestamoISO',
      'FechaPrestamo',
      'Estado',
      'NotasPrestamo',
      'FechaDevolucionISO',
      'FechaDevolucion',
      'CondicionDevolucion',
      'NotasDevolucion'
    ]
  ];

  state.loans.forEach(item => {
    rows.push([
      item.id,
      item.personType,
      item.personName,
      item.itemCategory || '',
      item.itemName,
      item.quantity,
      item.loanDateISO,
      item.loanDateLabel,
      item.status,
      item.loanNotes,
      item.returnDateISO,
      item.returnDateLabel,
      item.returnCondition,
      item.returnNotes
    ]);
  });

  const csv = rows.map(row => row.map(csvEscape).join(',')).join('\n');
  const fileName = `prestamos-musicala-${getTodayFileSafe()}.csv`;
  downloadFile(csv, fileName, 'text/csv;charset=utf-8;');
}

function handleImportJSON(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const text = String(e.target?.result || '');
      const parsed = JSON.parse(text);

      let importedCatalogs = null;
      let importedLoans = null;

      if (Array.isArray(parsed)) {
        importedLoans = parsed;
      } else if (parsed && typeof parsed === 'object') {
        importedLoans = Array.isArray(parsed.loans) ? parsed.loans : [];
        importedCatalogs = parsed.catalogs && typeof parsed.catalogs === 'object'
          ? normalizeCatalogs(parsed.catalogs)
          : null;
      }

      if (!Array.isArray(importedLoans)) {
        alert('El archivo JSON no tiene un formato válido.');
        resetImportInput();
        return;
      }

      const confirmed = window.confirm(
        '¿Quieres reemplazar los registros actuales con los del archivo importado?'
      );

      if (!confirmed) {
        resetImportInput();
        return;
      }

      state.loans = importedLoans.map(normalizeImportedLoan).filter(Boolean);

      if (importedCatalogs) {
        state.catalogs = importedCatalogs;
      } else {
        rebuildCatalogsFromLoansIfUseful();
      }

      persistLoans();
      persistCatalogs();
      renderAll();
      alert('Respaldo importado correctamente.');
    } catch (error) {
      console.error('Error importando JSON:', error);
      alert('No se pudo importar el archivo JSON.');
    } finally {
      resetImportInput();
    }
  };

  reader.onerror = () => {
    alert('No se pudo leer el archivo.');
    resetImportInput();
  };

  reader.readAsText(file);
}

function resetImportInput() {
  if (importFile) importFile.value = '';
}

function rebuildCatalogsFromLoansIfUseful() {
  const peopleMap = new Map();
  const categoriesMap = new Map();
  const itemsMap = new Map();

  state.loans.forEach(loan => {
    if (loan.personName && loan.personType) {
      const key = `${loan.personType}__${normalizeText(loan.personName)}`;
      if (!peopleMap.has(key)) {
        peopleMap.set(key, {
          id: generateCatalogId('PER'),
          type: loan.personType,
          name: loan.personName
        });
      }
    }

    if (loan.itemCategory) {
      const catKey = normalizeText(loan.itemCategory);
      if (!categoriesMap.has(catKey)) {
        categoriesMap.set(catKey, {
          id: generateCatalogId('CAT'),
          name: loan.itemCategory
        });
      }
    }

    if (loan.itemName && loan.itemCategory) {
      const itemKey = `${normalizeText(loan.itemCategory)}__${normalizeText(loan.itemName)}`;
      if (!itemsMap.has(itemKey)) {
        itemsMap.set(itemKey, {
          id: generateCatalogId('ITE'),
          category: loan.itemCategory,
          name: loan.itemName
        });
      }
    }
  });

  state.catalogs = {
    people: Array.from(peopleMap.values()),
    categories: Array.from(categoriesMap.values()),
    items: Array.from(itemsMap.values())
  };
}

function normalizeImportedLoan(item) {
  if (!item || typeof item !== 'object') return null;

  const normalizedCategory =
    String(item.itemCategory || item.category || '').trim();

  return {
    id: String(item.id || generateLoanId(new Date())),
    personType: String(item.personType || ''),
    personName: String(item.personName || ''),
    itemCategory: normalizedCategory,
    itemName: String(item.itemName || item.elemento || ''),
    quantity: Number(item.quantity || 1),
    loanDateISO: String(item.loanDateISO || item.createdAt || new Date().toISOString()),
    loanDateLabel: String(
      item.loanDateLabel ||
      formatDateTime(new Date(item.loanDateISO || item.createdAt || Date.now()))
    ),
    status: normalizeLoanStatus(String(item.status || 'Prestado')),
    loanNotes: String(item.loanNotes || ''),
    returnDateISO: String(item.returnDateISO || ''),
    returnDateLabel: String(item.returnDateLabel || ''),
    returnCondition: String(item.returnCondition || ''),
    returnNotes: String(item.returnNotes || ''),
    createdAt: String(item.createdAt || item.loanDateISO || new Date().toISOString()),
    updatedAt: String(item.updatedAt || item.loanDateISO || new Date().toISOString())
  };
}

function normalizeCatalogs(raw) {
  const peopleRaw = Array.isArray(raw.people) ? raw.people : [];
  const categoriesRaw = Array.isArray(raw.categories) ? raw.categories : [];
  const itemsRaw = Array.isArray(raw.items) ? raw.items : [];

  return {
    people: dedupeCatalogPeople(
      peopleRaw
        .map(item => ({
          id: String(item.id || generateCatalogId('PER')),
          type: String(item.type || '').trim(),
          name: String(item.name || '').trim()
        }))
        .filter(item => item.type && item.name)
    ),
    categories: dedupeCatalogCategories(
      categoriesRaw
        .map(item => ({
          id: String(item.id || generateCatalogId('CAT')),
          name: String(item.name || '').trim()
        }))
        .filter(item => item.name)
    ),
    items: dedupeCatalogItems(
      itemsRaw
        .map(item => ({
          id: String(item.id || generateCatalogId('ITE')),
          category: String(item.category || '').trim(),
          name: String(item.name || '').trim()
        }))
        .filter(item => item.category && item.name)
    )
  };
}

/* =========================
   DOWNLOAD
========================= */
function downloadFile(content, fileName, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);

  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* =========================
   HELPERS DE UI
========================= */
function resetLoanForm() {
  loanForm?.reset();
  if (quantity) quantity.value = '1';
  populatePersonSelect(personType?.value || '', '');
  populateCategorySelect('');
  populateItemSelect('', '');
}

function handleClearForm() {
  if (quantity) quantity.value = '1';
  populatePersonSelect(personType?.value || '', '');
  populateCategorySelect('');
  populateItemSelect('', '');
}

/* =========================
   UTILIDADES
========================= */
function generateLoanId(date = new Date()) {
  const part1 = date.getTime().toString(36).toUpperCase();
  const part2 = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `PRES-${part1}-${part2}`;
}

function generateCatalogId(prefix = 'CAT') {
  return `${prefix}-${Math.random().toString(36).slice(2, 10).toUpperCase()}`;
}

function formatDateTime(dateInput) {
  const date = dateInput instanceof Date ? dateInput : new Date(dateInput);

  if (Number.isNaN(date.getTime())) return 'Fecha inválida';

  return new Intl.DateTimeFormat('es-CO', {
    dateStyle: 'short',
    timeStyle: 'short'
  }).format(date);
}

function getTodayFileSafe() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function csvEscape(value) {
  const safe = String(value ?? '');
  if (safe.includes(',') || safe.includes('"') || safe.includes('\n')) {
    return `"${safe.replace(/"/g, '""')}"`;
  }
  return safe;
}

function escapeHTML(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function buildNotesSummary(loan) {
  const chunks = [];

  if (loan.loanNotes) chunks.push(`Préstamo: ${loan.loanNotes}`);
  if (loan.returnNotes) chunks.push(`Devolución: ${loan.returnNotes}`);

  return chunks.length ? chunks.join(' | ') : 'Sin notas';
}

function renderStatusBadgeText(status) {
  const safeStatus = normalizeLoanStatus(status);
  if (safeStatus === 'Con novedad') return 'Con novedad';
  if (safeStatus === 'Devuelto') return 'Devuelto';
  return 'Prestado';
}

function normalizeLoanStatus(status) {
  const safe = String(status || '').trim().toLowerCase();
  if (safe === 'con novedad') return 'Con novedad';
  if (safe === 'devuelto') return 'Devuelto';
  return 'Prestado';
}

function getStatusClass(status) {
  const safeStatus = normalizeLoanStatus(status);
  if (safeStatus === 'Con novedad') return 'status-novedad';
  if (safeStatus === 'Devuelto') return 'status-devuelto';
  return 'status-prestado';
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function comparePeople(a, b) {
  if (a.type !== b.type) {
    return a.type.localeCompare(b.type, 'es', { sensitivity: 'base' });
  }
  return a.name.localeCompare(b.name, 'es', { sensitivity: 'base' });
}

function compareItems(a, b) {
  if (a.category !== b.category) {
    return a.category.localeCompare(b.category, 'es', { sensitivity: 'base' });
  }
  return a.name.localeCompare(b.name, 'es', { sensitivity: 'base' });
}

function dedupeCatalogPeople(list) {
  const map = new Map();
  list.forEach(item => {
    const key = `${item.type}__${normalizeText(item.name)}`;
    if (!map.has(key)) map.set(key, item);
  });
  return Array.from(map.values());
}

function dedupeCatalogCategories(list) {
  const map = new Map();
  list.forEach(item => {
    const key = normalizeText(item.name);
    if (!map.has(key)) map.set(key, item);
  });
  return Array.from(map.values());
}

function dedupeCatalogItems(list) {
  const map = new Map();
  list.forEach(item => {
    const key = `${normalizeText(item.category)}__${normalizeText(item.name)}`;
    if (!map.has(key)) map.set(key, item);
  });
  return Array.from(map.values());
}
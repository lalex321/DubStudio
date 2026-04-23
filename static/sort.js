// Универсальная сортировка HTML-таблицы по клику на <th class="sortable">.
// Для числовых колонок берёт значение из data-атрибута активной метрики
// (если у ячейки есть data-{metric}), иначе — textContent.
// Для текстовых — textContent, либо value у вложенного <input>.
(function () {
  function cellValue(cell, type, metric) {
    if (type === 'num') {
      const dv = metric && cell.dataset[metric];
      const raw = dv !== undefined && dv !== '' ? dv : cell.textContent.trim();
      const n = Number(raw);
      return isNaN(n) ? 0 : n;
    }
    if (cell.dataset.sortValue !== undefined) return cell.dataset.sortValue.toLowerCase();
    const inp = cell.querySelector('input');
    const text = inp ? inp.value : cell.textContent;
    return (text || '').trim().toLowerCase();
  }

  function sortTable(table, col, dir) {
    const tbody = table.tBodies[0];
    if (!tbody) return;
    const ths = table.tHead.rows[0].cells;
    const type = ths[col].dataset.sortType || 'text';
    const metric = table.dataset.metric || '';
    const rows = Array.from(tbody.rows);
    rows.sort((a, b) => {
      const va = cellValue(a.cells[col], type, metric);
      const vb = cellValue(b.cells[col], type, metric);
      if (va < vb) return dir === 'asc' ? -1 : 1;
      if (va > vb) return dir === 'asc' ? 1 : -1;
      return 0;
    });
    for (const r of rows) tbody.appendChild(r);
  }

  function updateArrows(table, activeCol, dir) {
    const ths = table.tHead.rows[0].cells;
    for (let i = 0; i < ths.length; i++) {
      const arrow = ths[i].querySelector('.sort-arrow');
      if (!arrow) continue;
      arrow.textContent = i === activeCol ? (dir === 'asc' ? '↑' : '↓') : '';
    }
  }

  window.attachSort = function (table) {
    if (!table) return;
    const ths = table.tHead.rows[0].cells;
    // initial state: если у th есть data-sort-dir, считаем эту колонку активной
    let activeCol = -1;
    let activeDir = 'asc';
    for (let i = 0; i < ths.length; i++) {
      if (ths[i].dataset.sortDir) {
        activeCol = i;
        activeDir = ths[i].dataset.sortDir;
      }
    }
    const resort = () => {
      if (activeCol >= 0) {
        sortTable(table, activeCol, activeDir);
        updateArrows(table, activeCol, activeDir);
      }
    };
    table._resort = resort;  // чтобы внешний код (смена метрики) мог вызвать
    resort();

    for (let i = 0; i < ths.length; i++) {
      const th = ths[i];
      if (!th.classList.contains('sortable')) continue;
      th.addEventListener('click', () => {
        if (activeCol === i) {
          activeDir = activeDir === 'asc' ? 'desc' : 'asc';
        } else {
          activeCol = i;
          activeDir = 'asc';
        }
        sortTable(table, activeCol, activeDir);
        updateArrows(table, activeCol, activeDir);
      });
    }
  };
})();

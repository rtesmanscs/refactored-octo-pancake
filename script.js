let bom, pack, anc;

// ===== Utilities =====
const qs = (s, el=document) => el.querySelector(s);
const qsa = (s, el=document) => Array.from(el.querySelectorAll(s));
const MASS_TOL_PCT = 5;

// Unit conversions
function toKg(value, unit) {
  const v = parseFloat(value);
  if (isNaN(v)) return null;
  switch (unit) {
    case 'kg': return v;
    case 'g':  return v / 1000;
    case 'lb': return v * 0.45359237;
    case 't':  return v * 1000;
    default:   return null;
  }
}
function toM2(value, unit) {
  const v = parseFloat(value);
  if (isNaN(v)) return null;
  switch (unit) {
    case 'm2': return v;
    case 'ft2': return v * 0.09290304;
    case 'yd2': return v * 0.83612736;
    default:     return null;
  }
}
function toKm(value, unit) {
  const v = parseFloat(value);
  if (isNaN(v)) return null;
  switch (unit) {
    case 'km': return v;
    case 'mi': return v * 1.60934;
    default:     return null;
  }
}
function normalizeQtyUnit(value, unit) {
  const v = parseFloat(value);
  if (isNaN(v)) return { std_value: null, std_unit: unit || null, value: value, unit };
  if (['kg','g','lb','t'].includes(unit))  return { std_value: toKg(v, unit),  std_unit: 'kg',  value: v, unit };
  if (['m2','ft2','yd2'].includes(unit))   return { std_value: toM2(v, unit),  std_unit: 'm2', value: v, unit };
  return { std_value: v, std_unit: unit, value: v, unit };
}

// CSV helper
function toCSV(rows, columns) {
  const esc = (val) => {
    if (val === null || val === undefined) return '';
    const s = String(val);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const header = columns.map(esc).join(',');
  const lines  = rows.map(r => columns.map(c => esc(r[c])).join(','));
  return [header, ...lines].join('\n');
}

// ===== WAIT FOR DOM TO LOAD =====
document.addEventListener('DOMContentLoaded', function() {
  console.log('DOM loaded, initializing form...');
  initializeApp();
});

function initializeApp() {
  // ===== Form refs =====
  const form   = qs('#lcaForm');
  const errors = qs('#errors');
  const okMsg  = qs('#okMsg');

  if (!form) {
    console.error('Form not found!');
    return;
  }

  // ===== Basis toggles =====
  const basisTypeSel = '#basis_type';
  const basisBlocks = {
    mass:  qs('#basis_mass'),
    area:  qs('#basis_area'),
    count: qs('#basis_count')
  };
  
  function showBasis(which) {
    Object.values(basisBlocks).forEach(el => el && el.classList.add('hidden'));
    if (which && basisBlocks[which]) basisBlocks[which].classList.remove('hidden');
    updateProductMassQC();
  }
  
  const basisTypeEl = qs(basisTypeSel);
  if (basisTypeEl) {
    basisTypeEl.addEventListener('change', e => showBasis(e.target.value));
  }

  // ===== Production toggles & derived =====
  const prodModeInputs   = qsa('input[name="prod_mode"]');
  const wrapFacilityShare = qs('#prod_facility_share');
  const wrapProductDirect = qs('#prod_product_direct');
  
  function updateProdMode() {
    const mode = (prodModeInputs.find(r => r.checked) || {}).value;
    if (wrapFacilityShare) wrapFacilityShare.classList.toggle('hidden', mode !== 'facility_share');
    if (wrapProductDirect) wrapProductDirect.classList.toggle('hidden', mode !== 'product_direct');
  }
  prodModeInputs.forEach(r => r.addEventListener('change', updateProdMode));

  const facilityQty  = qs('input[name="facility_total_qty"]');
  const facilityUnit = qs('[name="facility_total_unit"]');
  const sharePct     = qs('input[name="product_share_pct"]');
  const derivedQty   = qs('input[name="product_qty_derived"]');
  const derivedUnit  = qs('input[name="product_unit_derived"]');
  
  function recalcDerived() {
    if (!facilityQty || !sharePct || !derivedQty) return;
    const f = parseFloat(facilityQty.value || '0');
    const p = parseFloat(sharePct.value   || '0');
    derivedQty.value  = (!isNaN(f) && !isNaN(p) && f >= 0 && p >= 0) ? (f * (p/100)).toFixed(6) : '';
    derivedUnit.value = (facilityUnit?.value) || '';
  }
  [facilityQty, sharePct, facilityUnit].filter(Boolean).forEach(i => i.addEventListener('input', recalcDerived));

  // ===== Inline Transport Section (BOM / PACK / ANC) =====
  const MODE_KEYS = ['road','rail','ship','air'];

  function makeInlineTransportSection(opts) {
    const {
      sectionKey, headEl, bodyEl, compactToggleEl,
      addRowBtnEl, statusEl,
      basePlaceholders,
      dualMass = false
    } = opts;

    if (!headEl || !bodyEl || !compactToggleEl || !addRowBtnEl) {
      console.error(`Missing elements for section ${sectionKey}`);
      return null;
    }

    const selectedModes = new Set();
    const modeUnits = {};

    qsa(`[data-mode$="_${sectionKey}"]`).forEach(cb => {
      const key = cb.dataset.mode;
      cb.addEventListener('change', (e) => {
        if (e.target.checked) selectedModes.add(key); else selectedModes.delete(key);
        renderHeader(); syncRowsToModes();
      });
    });
    
    qsa(`[data-mode-unit$="_${sectionKey}"]`).forEach(sel => {
      const key = sel.dataset.modeUnit;
      modeUnits[key] = sel.value || 'km';
      sel.addEventListener('change', (e) => { modeUnits[key] = e.target.value; renderHeader(); });
    });

    compactToggleEl.addEventListener('change', syncRowsToModes);

    function renderHeader() {
      const showCompact = compactToggleEl.checked;
      const firstRow = document.createElement('tr');

      let leftColsHTML = '';
      if (dualMass) {
        leftColsHTML = `
          <th rowspan="2">Component</th>
          <th colspan="2" class="num">Input mass (per product)</th>
          <th colspan="2" class="num">Used in final product</th>
        `;
      } else {
        leftColsHTML = `
          <th rowspan="2">Component</th>
          <th colspan="2" class="num">Mass per product</th>
        `;
      }

      firstRow.innerHTML = `
        ${leftColsHTML}
        ${showCompact ? '' : '<th rowspan="2">Supplier</th><th rowspan="2">Country</th><th rowspan="2" class="num">Pre-%</th><th rowspan="2" class="num">Post-%</th>'}
        ${selectedModes.size ? `<th colspan="${selectedModes.size}" class="num">Inbound transport distance</th>` : ''}
        <th rowspan="2" class="shrink"></th>
      `;

      const secondRow = document.createElement('tr');
      const subCols = dualMass ? ['Input value','Input unit','Used value','Used unit'] : ['Value','Unit'];
      subCols.forEach(lbl => {
        const th = document.createElement('th'); th.className = 'num'; th.textContent = lbl;
        secondRow.appendChild(th);
      });

      if (selectedModes.size) {
        Array.from(selectedModes).forEach(fullKey => {
          const short = fullKey.split('_')[0];
          const unit  = modeUnits[fullKey] || 'km';
          const label = short[0].toUpperCase() + short.slice(1);
          const th = document.createElement('th'); th.className='num';
          th.textContent = `${label} (${unit})`;
          secondRow.appendChild(th);
        });
      }

      headEl.innerHTML = '';
      headEl.appendChild(firstRow);
      headEl.appendChild(secondRow);
    }

    function addRow() {
      console.log(`Adding row to ${sectionKey}...`);
      const tr = document.createElement('tr');

      let massCellsHTML = '';
      if (dualMass) {
        massCellsHTML = `
          <td class="num"><input name="${sectionKey}_mass_input_value[]" type="number" min="0" step="0.0001" required class="col-small" placeholder="0.0000"></td>
          <td class="num">
            <select name="${sectionKey}_mass_input_unit[]" class="col-compact">
              <option value="kg">kg</option><option value="g">g</option><option value="lb">lb</option><option value="t">t</option>
            </select>
          </td>
          <td class="num"><input name="${sectionKey}_mass_used_value[]" type="number" min="0" step="0.0001" required class="col-small" placeholder="0.0000"></td>
          <td class="num">
            <select name="${sectionKey}_mass_used_unit[]" class="col-compact">
              <option value="kg">kg</option><option value="g">g</option><option value="lb">lb</option><option value="t">t</option>
            </select>
          </td>
        `;
      } else {
        massCellsHTML = `
          <td class="num"><input name="${sectionKey}_mass_value[]" type="number" min="0" step="0.0001" required class="col-small" placeholder="0.0000"></td>
          <td class="num">
            <select name="${sectionKey}_mass_unit[]" class="col-compact">
              <option value="kg">kg</option><option value="g">g</option><option value="lb">lb</option><option value="t">t</option>
            </select>
          </td>
        `;
      }

      tr.innerHTML = `
        <td><input name="${sectionKey}_component[]" required placeholder="${basePlaceholders.component}"></td>
        ${massCellsHTML}
        ${compactToggleEl.checked ? '' : `
          <td><input name="${sectionKey}_supplier[]" placeholder="${basePlaceholders.supplier}"></td>
          <td><input name="${sectionKey}_country[]" placeholder="Country (optional)"></td>
          <td class="num"><input name="${sectionKey}_pre_pct[]" type="number" min="0" max="100" step="0.1" class="col-compact" placeholder="0.0"></td>
          <td class="num"><input name="${sectionKey}_post_pct[]" type="number" min="0" max="100" step="0.1" class="col-compact" placeholder="0.0"></td>
        `}
        <td class="shrink"><button type="button" class="remove">Remove</button></td>
      `;

      const lastTd = tr.lastElementChild;
      Array.from(selectedModes).forEach(fullKey => {
        const td = document.createElement('td'); td.className='num'; td.setAttribute('data-mode-cell', fullKey);
        td.innerHTML = `<input name="dist_${fullKey}[]" type="number" min="0" step="0.01" value="0" class="col-trans">`;
        tr.insertBefore(td, lastTd);
      });

      tr.querySelector('.remove').addEventListener('click', () => { 
        tr.remove(); 
        if (sectionKey==='bom' || sectionKey==='pack') updateProductMassQC(); 
      });
      
      if (sectionKey==='bom' || sectionKey==='pack') {
        tr.querySelectorAll('input, select').forEach(inp => inp.addEventListener('input', updateProductMassQC));
      }

      bodyEl.appendChild(tr);
      if (sectionKey==='bom' || sectionKey==='pack') updateProductMassQC();
    }

    function syncRowsToModes() {
      const rows = qsa(`#${bodyEl.id} tr`);
      rows.forEach(tr => {
        const lastTd = tr.lastElementChild;
        Array.from(selectedModes).forEach(fullKey => {
          if (!tr.querySelector(`[data-mode-cell="${fullKey}"]`)) {
            const td = document.createElement('td'); td.className='num'; td.setAttribute('data-mode-cell', fullKey);
            td.innerHTML = `<input name="dist_${fullKey}[]" type="number" min="0" step="0.01" value="0" class="col-trans">`;
            tr.insertBefore(td, lastTd);
          }
        });
        MODE_KEYS.map(k => `${k}_${sectionKey}`).forEach(fullKey => {
          if (!selectedModes.has(fullKey)) {
            const cell = tr.querySelector(`[data-mode-cell="${fullKey}"]`);
            if (cell) cell.remove();
          }
        });

        const showCompact = compactToggleEl.checked;
        const hasSupplier = tr.querySelector(`[name="${sectionKey}_supplier[]"]`);
        if (showCompact && hasSupplier) {
          [`${sectionKey}_supplier[]`,`${sectionKey}_country[]`,`${sectionKey}_pre_pct[]`,`${sectionKey}_post_pct[]`].forEach(name => {
            const el = tr.querySelector(`[name="${name}"]`);
            if (el) el.closest('td').remove();
          });
        } else if (!showCompact && !hasSupplier) {
          const firstModeCell = tr.querySelector('[data-mode-cell]') || tr.lastElementChild;
          const supplierTd = document.createElement('td'); supplierTd.innerHTML = `<input name="${sectionKey}_supplier[]" placeholder="${basePlaceholders.supplier}">`;
          const countryTd = document.createElement('td'); countryTd.innerHTML = `<input name="${sectionKey}_country[]" placeholder="Country (optional)">`;
          const preTd = document.createElement('td'); preTd.className='num'; preTd.innerHTML = `<input name="${sectionKey}_pre_pct[]" type="number" min="0" max="100" step="0.1" class="col-compact" placeholder="0.0">`;
          const postTd = document.createElement('td'); postTd.className='num'; postTd.innerHTML = `<input name="${sectionKey}_post_pct[]" type="number" min="0" max="100" step="0.1" class="col-compact" placeholder="0.0">`;
          tr.insertBefore(supplierTd, firstModeCell);
          tr.insertBefore(countryTd, firstModeCell);
          tr.insertBefore(preTd, firstModeCell);
          tr.insertBefore(postTd, firstModeCell);
        }
      });
      renderHeader();
    }

    addRowBtnEl.addEventListener('click', (e) => {
      e.preventDefault();
      addRow();
    });

    renderHeader();
    addRow();

    return { collectRows, statusEl, addRow, renderHeader, syncRowsToModes };

    function collectRows() {
      const rows = qsa(`#${bodyEl.id} tr`);
      const items = [];
      const inbound_rows = [];
      const inbound_flat = [];

      rows.forEach(tr => {
        const component = tr.querySelector(`[name="${sectionKey}_component[]"]`)?.value?.trim();
        let massKg = null, massKgUsed = null;

        if (dualMass) {
          const iv = tr.querySelector(`[name="${sectionKey}_mass_input_value[]"]`)?.value;
          const iu = tr.querySelector(`[name="${sectionKey}_mass_input_unit[]"]`)?.value || 'kg';
          const uv = tr.querySelector(`[name="${sectionKey}_mass_used_value[]"]`)?.value;
          const uu = tr.querySelector(`[name="${sectionKey}_mass_used_unit[]"]`)?.value || 'kg';
          massKg     = toKg(iv, iu);
          massKgUsed = toKg(uv, uu);
        } else {
          const v = tr.querySelector(`[name="${sectionKey}_mass_value[]"]`)?.value;
          const u = tr.querySelector(`[name="${sectionKey}_mass_unit[]"]`)?.value || 'kg';
          massKg     = toKg(v, u);
          massKgUsed = massKg;
        }

        if (!(component && massKgUsed && massKgUsed > 0)) return;

        const supplierEl = tr.querySelector(`[name="${sectionKey}_supplier[]"]`);
        const countryEl  = tr.querySelector(`[name="${sectionKey}_country[]"]`);
        const preEl      = tr.querySelector(`[name="${sectionKey}_pre_pct[]"]`);
        const postEl     = tr.querySelector(`[name="${sectionKey}_post_pct[]"]`);
        const supplier   = supplierEl ? supplierEl.value?.trim() : null;
        const country    = countryEl  ? countryEl.value?.trim()  : null;
        const pre        = preEl  && preEl.value  !== '' ? parseFloat(preEl.value)  : null;
        const post       = postEl && postEl.value !== '' ? parseFloat(postEl.value) : null;

        const distances = {};
        Array.from(selectedModes).forEach(fullKey => {
          const inp  = tr.querySelector(`[data-mode-cell="${fullKey}"] input`);
          const val  = inp ? parseFloat(inp.value || '0') : 0;
          const unit = modeUnits[fullKey] || 'km';
          const km   = toKm(val, unit) || 0;
          distances[fullKey] = { value: isNaN(val)?0:val, unit, km };
          if (km > 0) {
            const short = fullKey.split('_')[0];
            inbound_flat.push({ component, mode: short, distance_value: val, distance_unit: unit, distance_km: km });
          }
        });

        items.push({
          component,
          mass_input_kg_per_product: massKg || 0,
          mass_used_kg_per_product:  massKgUsed || 0,
          supplier, country, pre_consumer_pct: pre, post_consumer_pct: post
        });
        inbound_rows.push({ component, distances });
      });

      const selectedShort = Array.from(selectedModes).map(fullKey => fullKey.split('_')[0]);
      const unitsShort = {};
      Array.from(selectedModes).forEach(fullKey => { unitsShort[fullKey.split('_')[0]] = modeUnits[fullKey]; });

      return { items, inbound_rows, inbound_flat, selected_modes: Array.from(new Set(selectedShort)), mode_units: unitsShort };
    }
  }

  // ===== Instantiate sections =====
  console.log('Creating BOM section...');
  bom = makeInlineTransportSection({
    sectionKey: 'bom',
    headEl: qs('#bomHead'), bodyEl: qs('#bomBody'),
    compactToggleEl: qs('#compactToggleBOM'),
    addRowBtnEl: qs('#addBom'), statusEl: qs('#bomStatus'),
    basePlaceholders: { component: 'e.g., PVC resin', supplier: 'Supplier (optional)' },
    dualMass: true
  });
  
  console.log('Creating Packaging section...');
  pack = makeInlineTransportSection({
    sectionKey: 'pack',
    headEl: qs('#packHead'), bodyEl: qs('#packBody'),
    compactToggleEl: qs('#compactTogglePACK'),
    addRowBtnEl: qs('#addPack'), statusEl: qs('#packStatus'),
    basePlaceholders: { component: 'e.g., Corrugated box (primary/secondary/tertiary)', supplier: 'Supplier (optional)' },
    dualMass: true
  });
  
  console.log('Creating Ancillary section...');
  anc = makeInlineTransportSection({
    sectionKey: 'anc',
    headEl: qs('#ancHead'), bodyEl: qs('#ancBody'),
    compactToggleEl: qs('#compactToggleANC'),
    addRowBtnEl: qs('#addAnc'), statusEl: qs('#ancStatus'),
    basePlaceholders: { component: 'e.g., Adhesive, ink, release liner', supplier: 'Supplier (optional)' }
  });

  // ===== Energy =====
  const energyList = qs('#energyList');
  function addEnergyEntry(data = {}) {
    if (!energyList) return;
    const wrapper = document.createElement('div'); wrapper.className = 'row';
    wrapper.innerHTML = `
      <label>Type
        <select name="energy_type[]" required>
          <option value="">Select</option>
          <option ${data.type==='electricity'?'selected':''} value="electricity">Electricity</option>
          <option ${data.type==='natural_gas'?'selected':''} value="natural_gas">Natural gas</option>
          <option ${data.type==='diesel'?'selected':''} value="diesel">Diesel</option>
          <option ${data.type==='propane'?'selected':''} value="propane">Propane</option>
          <option ${data.type==='steam'?'selected':''} value="steam">Steam</option>
          <option ${data.type==='other'?'selected':''} value="other">Other</option>
        </select>
      </label>
      <label>Amount<input name="energy_amount[]" type="number" min="0" step="0.0001" required></label>
      <label>Unit
        <select name="energy_unit[]" required>
          <option value="kWh">kWh</option>
          <option value="therm">therm</option>
          <option value="m3">m³</option>
          <option value="L">L</option>
          <option value="kg">kg</option>
          <option value="MJ">MJ</option>
        </select>
      </label>
      <label>Notes<input name="energy_notes[]" placeholder="e.g., market-based, supplier mix"></label>
      <button type="button" class="remove">Remove</button>
    `;
    wrapper.querySelector('.remove').addEventListener('click', () => wrapper.remove());
    energyList.appendChild(wrapper);
  }
  
  const addEnergyBtn = qs('#addEnergy');
  if (addEnergyBtn) {
    addEnergyBtn.addEventListener('click', (e) => { 
      e.preventDefault(); 
      console.log('Adding energy entry...');
      addEnergyEntry(); 
    });
  }
  addEnergyEntry();

  // ===== Product-mass QC =====
  function sumUsedMassKgFromSection(items) {
    return (items || []).reduce((s, it) => s + (parseFloat(it.mass_used_kg_per_product) || 0), 0);
  }
  
  function getProductBasisMassKg() {
    const bt = qs('#basis_type')?.value;
    if (bt === 'mass') {
      return toKg(form.product_mass_value?.value, form.product_mass_unit?.value) || null;
    }
    if (bt === 'area') {
      return toKg(form.unit_weight_value_area?.value, form.unit_weight_unit_area?.value) || null;
    }
    if (bt === 'count') {
      return toKg(form.unit_weight_value_count?.value, form.unit_weight_unit_count?.value) || null;
    }
    return null;
  }
  
  function updateProductMassQC() {
    if (!bom || !pack) return;
    const statusEl = qs('#bomStatus');
    if (!statusEl) return;
    
    const bomData  = bom.collectRows();
    const packData = pack.collectRows();
    const usedBom  = sumUsedMassKgFromSection(bomData.items);
    const usedPack = sumUsedMassKgFromSection(packData.items);
    const usedTotal = usedBom + usedPack;

    const context = `Used mass — BOM: ${usedBom.toFixed(4)} kg, Packaging: ${usedPack.toFixed(4)} kg, Total: ${usedTotal.toFixed(4)} kg. `;
    const productMass = getProductBasisMassKg();

    if (!productMass || productMass <= 0) {
      statusEl.className = 'muted';
      statusEl.textContent = context + 'Enter product basis mass (or unit weight) to enable QC.';
      return;
    }

    const diff = usedTotal - productMass;
    const pct  = (diff / productMass) * 100;
    const withinTol = Math.abs(pct) <= MASS_TOL_PCT;
    statusEl.className = withinTol ? 'ok' : 'warn';
    statusEl.textContent = context +
      `Product mass: ${productMass.toFixed(4)} kg | Δ = ${diff.toFixed(4)} kg (${pct.toFixed(2)}%). ` +
      (withinTol ? `Within tolerance (±${MASS_TOL_PCT}%).` : 'Outside tolerance—adjust or explain.');
  }
  
  qsa('input[name="product_mass_value"], select[name="product_mass_unit"], \
       input[name="unit_weight_value_area"], select[name="unit_weight_unit_area"], \
       input[name="unit_weight_value_count"], select[name="unit_weight_unit_count"]').forEach(el => {
    el.addEventListener('input', updateProductMassQC);
    el.addEventListener('change', updateProductMassQC);
  });

  // ===== Build JSON payload =====
  function buildExportJSON() {
    const meta = {
      company: form.company?.value?.trim(),
      email:   form.email?.value?.trim(),
      facility: form.facility?.value?.trim(),
      period:  form.period?.value?.trim(),
      grid_region: form.grid_region?.value?.trim(),
      confidentiality: form.confidentiality?.value
    };
    const product = { name: form.product?.value?.trim(), sku: form.sku?.value?.trim() || null };

    const bt = qs(basisTypeSel)?.value;
    const basis = { type: bt, mass_kg: null, area_m2: null, units: null, unit_desc: null, unit_weight_kg: null };
    if (bt === 'mass') basis.mass_kg = toKg(form.product_mass_value?.value, form.product_mass_unit?.value);
    if (bt === 'area') {
      basis.area_m2       = toM2(form.product_area_value?.value, form.product_area_unit?.value);
      basis.unit_weight_kg = toKg(form.unit_weight_value_area?.value, form.unit_weight_unit_area?.value);
    }
    if (bt === 'count') {
      basis.units        = parseFloat(form.product_units?.value || '0');
      basis.unit_desc    = form.product_unit_desc?.value?.trim();
      basis.unit_weight_kg = toKg(form.unit_weight_value_count?.value, form.unit_weight_unit_count?.value);
    }

    const mode = (qsa('input[name="prod_mode"]').find(r => r.checked) || {}).value || null;
    const production = {
      mode,
      facility_total_qty: form.facility_total_qty?.value ? parseFloat(form.facility_total_qty.value) : null,
      facility_total_unit: form.facility_total_unit?.value || null,
      product_share_pct: form.product_share_pct?.value ? parseFloat(form.product_share_pct.value) : null,
      product_qty_derived: form.product_qty_derived?.value ? parseFloat(form.product_qty_derived.value) : null,
      product_unit_derived: form.product_unit_derived?.value || null,
      product_direct_qty: form.product_direct_qty?.value ? parseFloat(form.product_direct_qty.value) : null,
      product_direct_unit: form.product_direct_unit?.value || null
    };
    const period_product_output = (mode === 'facility_share') ? production.product_qty_derived : production.product_direct_qty;

    const facilityStd = production.facility_total_qty != null && production.facility_total_unit
      ? normalizeQtyUnit(production.facility_total_qty, production.facility_total_unit)
      : null;
    const directStd = production.product_direct_qty != null && production.product_direct_unit
      ? normalizeQtyUnit(production.product_direct_qty, production.product_direct_unit)
      : null;

    const production_std = {
      facility_total_qty_std: facilityStd ? facilityStd.std_value : null,
      facility_total_unit_std: facilityStd ? facilityStd.std_unit : null,
      product_direct_qty_std: directStd ? directStd.std_value : null,
      product_direct_unit_std: directStd ? directStd.std_unit : null
    };

    const bomData  = bom ? bom.collectRows() : { items: [], inbound_rows: [], inbound_flat: [], selected_modes: [], mode_units: {} };
    const packData = pack ? pack.collectRows() : { items: [], inbound_rows: [], inbound_flat: [], selected_modes: [], mode_units: {} };
    const ancData  = anc ? anc.collectRows() : { items: [], inbound_rows: [], inbound_flat: [], selected_modes: [], mode_units: {} };

    const energy = qsa('#energyList .row').map(row => {
      const type   = row.querySelector('[name="energy_type[]"]')?.value;
      const amount = parseFloat(row.querySelector('[name="energy_amount[]"]')?.value || '0');
      const unit   = row.querySelector('[name="energy_unit[]"]')?.value;
      const notes  = row.querySelector('[name="energy_notes[]"]')?.value?.trim() || null;
      if (!type || !(amount >= 0)) return null;
      return { type, amount, unit, notes };
    }).filter(Boolean);

    return {
      meta,
      product,
      basis,
      production: { ...production, period_product_output },
      production_std,
      bom:        bomData.items,
      packaging:  packData.items,
      ancillary:  ancData.items,
      transport: {
        bom:        { selected_modes: bomData.selected_modes,   mode_units: bomData.mode_units,   inbound_rows: bomData.inbound_rows,   inbound_flat: bomData.inbound_flat },
        packaging:  { selected_modes: packData.selected_modes,  mode_units: packData.mode_units,  inbound_rows: packData.inbound_rows,  inbound_flat: packData.inbound_flat },
        ancillary:  { selected_modes: ancData.selected_modes,   mode_units: ancData.mode_units,   inbound_rows: ancData.inbound_rows,   inbound_flat: ancData.inbound_flat }
      },
      energy
    };
  }

  // ===== CSV/ZIP generation =====
  function computeTonKmRows(flatRows, items, moduleLabel, useInputMass = false) {
    const massByComp = {};
    items.forEach(it => {
      // For A2 transport: use INPUT mass (raw material being transported)
      // For A3 materials in product: use USED mass
      let mass;
      if (useInputMass && typeof it.mass_input_kg_per_product === 'number') {
        mass = it.mass_input_kg_per_product;
      } else if (typeof it.mass_used_kg_per_product === 'number') {
        mass = it.mass_used_kg_per_product;
      } else {
        mass = it.mass_kg_per_product; // ancillary fallback
      }
      massByComp[it.component] = mass;
    });
    return flatRows.map(r => {
      const mass = massByComp[r.component] || 0;
      const ton_km = (r.distance_km || 0) * (mass / 1000);
      return {
        component: r.component,
        mode: r.mode,
        distance_value: r.distance_value,
        distance_unit: r.distance_unit,
        distance_km: r.distance_km,
        mass_kg_per_product: mass,
        ton_km_per_product: +ton_km.toFixed(6),
        module: moduleLabel
      };
    });
  }

  function energyRowsToCSV(energy, period_product_output) {
    const rows = energy.map(e => {
      const per = (period_product_output && period_product_output > 0) ? (e.amount / period_product_output) : null;
      return {
        type: e.type, amount: e.amount, unit: e.unit,
        per_product_amount: per !== null ? +per.toFixed(8) : null,
        per_product_note: per !== null ? 'Computed as amount / period_product_output' : null,
        notes: e.notes || null
      };
    });
    const cols = ["type","amount","unit","per_product_amount","per_product_note","notes"];
    return toCSV(rows, cols);
  }

  function generateCSVs(payload) {
    const a1Rows = payload.bom.map(b => ({
      component: b.component,
      mass_input_kg_per_product: b.mass_input_kg_per_product,
      mass_used_kg_per_product:  b.mass_used_kg_per_product,
      supplier: b.supplier || null,
      country:  b.country  || null,
      pre_consumer_pct:  b.pre_consumer_pct,
      post_consumer_pct: b.post_consumer_pct,
      module: "A1"
    }));
    const a1Cols = ["component","mass_input_kg_per_product","mass_used_kg_per_product","supplier","country","pre_consumer_pct","post_consumer_pct","module"];
    const A1_materials_csv = toCSV(a1Rows, a1Cols);

    // A2: Use INPUT mass for transport (raw materials being shipped)
    const a2Rows = computeTonKmRows(payload.transport.bom.inbound_flat, payload.bom, "A2", true);
    const tCols  = ["component","mode","distance_value","distance_unit","distance_km","mass_kg_per_product","ton_km_per_product","module"];
    const A2_inbound_transport_bom_csv = toCSV(a2Rows, tCols);

    const a3_pack_rows = payload.packaging.map(p => ({
      component: p.component,
      mass_input_kg_per_product: p.mass_input_kg_per_product,
      mass_used_kg_per_product:  p.mass_used_kg_per_product,
      supplier: p.supplier || null,
      country:  p.country  || null,
      pre_consumer_pct:  p.pre_consumer_pct,
      post_consumer_pct: p.post_consumer_pct,
      module: "A3"
    }));
    const A3_packaging_materials_csv = toCSV(a3_pack_rows, a1Cols);

    // A3 Packaging transport: Use INPUT mass (raw packaging materials being shipped)
    const a3_pack_inbound_rows = computeTonKmRows(payload.transport.packaging.inbound_flat, payload.packaging, "A3", true);
    const A3_inbound_transport_packaging_csv = toCSV(a3_pack_inbound_rows, tCols);

    const a3_anc_rows = payload.ancillary.map(a => ({
      component: a.component,
      mass_kg_per_product: a.mass_kg_per_product,
      supplier: a.supplier || null,
      country:  a.country  || null,
      pre_consumer_pct:  a.pre_consumer_pct,
      post_consumer_pct: a.post_consumer_pct,
      module: "A3"
    }));
    const A3_ancillary_materials_csv = toCSV(a3_anc_rows, ["component","mass_kg_per_product","supplier","country","pre_consumer_pct","post_consumer_pct","module"]);

    const A3_inbound_transport_ancillary_csv = toCSV(
      computeTonKmRows(payload.transport.ancillary.inbound_flat, payload.ancillary, "A3"),
      tCols
    );

    const energy_csv = energyRowsToCSV(payload.energy, payload.production.period_product_output);

    const m = payload;
    const metaRow = [{
      company: m.meta.company,
      facility: m.meta.facility,
      period:   m.meta.period,
      product:  m.product.name,
      sku:      m.product.sku || '',
      basis_type:         m.basis.type,
      basis_mass_kg:      m.basis.mass_kg,
      basis_area_m2:      m.basis.area_m2,
      basis_units:        m.basis.units,
      basis_unit_desc:    m.basis.unit_desc,
      basis_unit_weight_kg: m.basis.unit_weight_kg,
      facility_total_qty_std:  m.production_std.facility_total_qty_std,
      facility_total_unit_std: m.production_std.facility_total_unit_std,
      product_direct_qty_std:  m.production_std.product_direct_qty_std,
      product_direct_unit_std: m.production_std.product_direct_unit_std,
      period_product_output:   m.production.period_product_output
    }];
    const metaCols = ["company","facility","period","product","sku","basis_type","basis_mass_kg","basis_area_m2","basis_units","basis_unit_desc","basis_unit_weight_kg","facility_total_qty_std","facility_total_unit_std","product_direct_qty_std","product_direct_unit_std","period_product_output"];
    const meta_csv = toCSV(metaRow, metaCols);

    return {
      "A1_materials.csv":                        A1_materials_csv,
      "A2_inbound_transport_bom.csv":            A2_inbound_transport_bom_csv,
      "A3_packaging_materials.csv":              A3_packaging_materials_csv,
      "A3_inbound_transport_packaging.csv":      A3_inbound_transport_packaging_csv,
      "A3_ancillary_materials.csv":              A3_ancillary_materials_csv,
      "A3_inbound_transport_ancillary.csv":      A3_inbound_transport_ancillary_csv,
      "energy.csv":                              energy_csv,
      "_meta_summary.csv":                       meta_csv,
      "README.txt": [
        "LCA/EPD Intake export",
        "- Materials mass are per product (kg); transport is ton-km per product (based on USED mass).",
        "- Energy includes per_product_amount if production output was provided.",
        "- Standardized production units (kg/m2) in _meta_summary.csv"
      ].join("\n")
    };
  }

  async function downloadZipOfCSVs(payload) {
    const csvs = generateCSVs(payload);
    const zip  = new JSZip();
    Object.entries(csvs).forEach(([name, content]) => zip.file(name, content));
    const blob = await zip.generateAsync({type: "blob"});
    saveAs(blob, "intake_export.zip");
  }

  // ===== Actions =====
  qs('#downloadJson').addEventListener('click', () => {
    const payload = buildExportJSON();
    const blob = new Blob([JSON.stringify(payload, null, 2)], {type: 'application/json'});
    saveAs(blob, 'lca_epd_intake.json');
    okMsg.textContent = 'JSON downloaded.'; setTimeout(()=>okMsg.textContent='', 2500);
  });
  
  qs('#downloadZip').addEventListener('click', async () => {
    const payload = buildExportJSON();
    await downloadZipOfCSVs(payload);
    okMsg.textContent = 'CSV ZIP downloaded.'; setTimeout(()=>okMsg.textContent='', 2500);
  });

  // ===== Submit validation =====
  form.addEventListener('submit', (e) => {
    errors.textContent = ''; okMsg.textContent = '';

    const hasAnyBOM = bom.collectRows().items.length > 0;
    if (!hasAnyBOM) {
      e.preventDefault(); errors.textContent = 'Add at least one BOM component with mass > 0.'; return;
    }

    const bt = qs(basisTypeSel).value;
    if (!bt) { e.preventDefault(); errors.textContent = 'Choose a Product Basis (mass, area, or count).'; return; }

    if (bt === 'mass') {
      const pm = toKg(form.product_mass_value.value, form.product_mass_unit.value);
      if (!(pm > 0)) { e.preventDefault(); errors.textContent = 'Provide product mass (value + unit).'; return; }
    }
    if (bt === 'area') {
      const v  = toM2(form.product_area_value.value, form.product_area_unit.value);
      const uw = toKg(form.unit_weight_value_area.value, form.unit_weight_unit_area.value);
      if (!(v  > 0)) { e.preventDefault(); errors.textContent = 'Provide product area (value + unit).'; return; }
      if (!(uw > 0)) { e.preventDefault(); errors.textContent = 'Provide weight per product unit (value + unit).'; return; }
    }
    if (bt === 'count') {
      const n  = parseFloat(form.product_units.value || '0');
      const uw = toKg(form.unit_weight_value_count.value, form.unit_weight_unit_count.value);
      if (!(n  > 0)) { e.preventDefault(); errors.textContent = 'Provide units per product (count).'; return; }
      if (!(uw > 0)) { e.preventDefault(); errors.textContent = 'Provide weight per product unit (value + unit).'; return; }
    }

    const mode = (qsa('input[name="prod_mode"]').find(r => r.checked) || {}).value;
    if (!mode) { e.preventDefault(); errors.textContent = 'Select how you want to report production data.'; return; }
    if (mode === 'facility_share') {
      const f = parseFloat(form.facility_total_qty.value);
      const p = parseFloat(form.product_share_pct.value);
      if (isNaN(f)) { e.preventDefault(); errors.textContent = 'Enter facility total production quantity.'; return; }
      if (!(p >= 0 && p <= 100)) { e.preventDefault(); errors.textContent = 'Product share must be 0–100%.'; return; }
    } else if (mode === 'product_direct') {
      const q = parseFloat(form.product_direct_qty.value);
      if (isNaN(q)) { e.preventDefault(); errors.textContent = 'Enter product production quantity for the period.'; return; }
    }

    const usedBom  = sumUsedMassKgFromSection(bom.collectRows().items);
    const usedPack = sumUsedMassKgFromSection(pack.collectRows().items);
    const usedTot  = usedBom + usedPack;
    const prodMass = getProductBasisMassKg();
    if (prodMass && prodMass > 0) {
      const diff = usedTot - prodMass;
      const pct  = (diff / prodMass) * 100;
      if (Math.abs(pct) > MASS_TOL_PCT) {
        e.preventDefault();
        errors.textContent = `Product mass vs Used(BOM+Packaging) mismatch (${pct.toFixed(2)}%). Must be within ±${MASS_TOL_PCT}%.`;
        return;
      }
    }

    okMsg.textContent = 'Form validated.'; setTimeout(()=>okMsg.textContent='', 2500);
    e.preventDefault();
  });

  // ===== XLSX Export =====
  (function addXlsxExport() {
    function ensureXLSX(callback) {
      if (window.XLSX) return callback();
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
      s.async = true;
      s.onload = callback;
      document.head.appendChild(s);
    }

    function aoaToSheetWithFormatting(aoa, numericFormats) {
      const XLSX = window.XLSX;
      const ws = XLSX.utils.aoa_to_sheet(aoa);

      const header = aoa[0] || [];
      const widths = header.map(h => {
        const w = Math.max(10, Math.min(30, String(h || '').length + 2));
        return { wch: w };
      });
      ws['!cols'] = widths;

      const range = XLSX.utils.decode_range(ws['!ref']);
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: 0, c: range.e.c }
      }) };

      ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft' };

      if (numericFormats && Object.keys(numericFormats).length) {
        const headerIndex = {};
        header.forEach((h, idx) => { headerIndex[String(h)] = idx; });

        Object.entries(numericFormats).forEach(([colHeader, fmt]) => {
          const c = headerIndex[colHeader];
          if (c == null) return;
          for (let r = 1; r <= range.e.r; r++) {
            const addr = window.XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            if (!cell) continue;
            const v = typeof cell.v === 'string' ? parseFloat(cell.v.replace(/,/g,'')) : cell.v;
            if (!isNaN(v)) {
              cell.t = 'n';
              cell.v = v;
              cell.z = fmt;
            }
          }
        });
      }

      return ws;
    }

    function addSheet(wb, name, rows, headers, numericFormats) {
      const aoa = [headers].concat(
        rows.map(r => headers.map(h => (r[h] === undefined ? null : r[h])))
      );
      const ws = aoaToSheetWithFormatting(aoa, numericFormats || {});
      window.XLSX.utils.book_append_sheet(wb, ws, name);
    }

    function buildWorkbookFromPayload(payload) {
      const XLSX = window.XLSX;
      const wb = XLSX.utils.book_new();

      const a1Rows = payload.bom.map(b => ({
        component: b.component,
        mass_input_kg_per_product: b.mass_input_kg_per_product,
        mass_used_kg_per_product:  b.mass_used_kg_per_product,
        supplier: b.supplier || null,
        country:  b.country  || null,
        pre_consumer_pct:  b.pre_consumer_pct,
        post_consumer_pct: b.post_consumer_pct,
        module: "A1"
      }));
      addSheet(
        wb,
        'A1_materials',
        a1Rows,
        ["component","mass_input_kg_per_product","mass_used_kg_per_product","supplier","country","pre_consumer_pct","post_consumer_pct","module"],
        { mass_input_kg_per_product: "0.0000", mass_used_kg_per_product: "0.0000", pre_consumer_pct: "0.0", post_consumer_pct: "0.0" }
      );

      const a2Rows = computeTonKmRows(payload.transport.bom.inbound_flat, payload.bom, "A2", true);
      addSheet(
        wb,
        'A2_inbound_transport_bom',
        a2Rows,
        ["component","mode","distance_value","distance_unit","distance_km","mass_kg_per_product","ton_km_per_product","module"],
        { distance_value: "0.00", distance_km: "0.00", mass_kg_per_product: "0.0000", ton_km_per_product: "0.000000" }
      );

      const a3PackRows = payload.packaging.map(p => ({
        component: p.component,
        mass_input_kg_per_product: p.mass_input_kg_per_product,
        mass_used_kg_per_product:  p.mass_used_kg_per_product,
        supplier: p.supplier || null,
        country:  p.country  || null,
        pre_consumer_pct:  p.pre_consumer_pct,
        post_consumer_pct: p.post_consumer_pct,
        module: "A3"
      }));
      addSheet(
        wb,
        'A3_packaging_materials',
        a3PackRows,
        ["component","mass_input_kg_per_product","mass_used_kg_per_product","supplier","country","pre_consumer_pct","post_consumer_pct","module"],
        { mass_input_kg_per_product: "0.0000", mass_used_kg_per_product: "0.0000", pre_consumer_pct: "0.0", post_consumer_pct: "0.0" }
      );
      
      const a3PackTransport = computeTonKmRows(payload.transport.packaging.inbound_flat, payload.packaging, "A3", true);
      addSheet(
        wb,
        'A3_inbound_transport_packaging',
        a3PackTransport,
        ["component","mode","distance_value","distance_unit","distance_km","mass_kg_per_product","ton_km_per_product","module"],
        { distance_value: "0.00", distance_km: "0.00", mass_kg_per_product: "0.0000", ton_km_per_product: "0.000000" }
      );

      const a3AncRows = payload.ancillary.map(a => ({
        component: a.component,
        mass_kg_per_product: a.mass_kg_per_product,
        supplier: a.supplier || null,
        country:  a.country  || null,
        pre_consumer_pct:  a.pre_consumer_pct,
        post_consumer_pct: a.post_consumer_pct,
        module: "A3"
      }));
      addSheet(
        wb,
        'A3_ancillary_materials',
        a3AncRows,
        ["component","mass_kg_per_product","supplier","country","pre_consumer_pct","post_consumer_pct","module"],
        { mass_kg_per_product: "0.0000", pre_consumer_pct: "0.0", post_consumer_pct: "0.0" }
      );
      
      const a3AncTransport = computeTonKmRows(payload.transport.ancillary.inbound_flat, payload.ancillary, "A3");
      addSheet(
        wb,
        'A3_inbound_transport_ancillary',
        a3AncTransport,
        ["component","mode","distance_value","distance_unit","distance_km","mass_kg_per_product","ton_km_per_product","module"],
        { distance_value: "0.00", distance_km: "0.00", mass_kg_per_product: "0.0000", ton_km_per_product: "0.000000" }
      );

      const energyRows = (payload.energy || []).map(e => {
        const ppo = payload.production?.period_product_output;
        const per = (ppo && ppo > 0) ? (e.amount / ppo) : null;
        return {
          type: e.type, amount: e.amount, unit: e.unit,
          per_product_amount: per !== null ? +per.toFixed(8) : null,
          per_product_note: per !== null ? 'Computed as amount / period_product_output' : null,
          notes: e.notes || null
        };
      });
      addSheet(
        wb,
        'energy',
        energyRows,
        ["type","amount","unit","per_product_amount","per_product_note","notes"],
        { amount: "0.0000", per_product_amount: "0.00000000" }
      );

      const m = payload;
      const metaRow = [{
        company: m.meta.company,
        facility: m.meta.facility,
        period: m.meta.period,
        product: m.product.name,
        sku: m.product.sku || '',
        basis_type: m.basis.type,
        basis_mass_kg: m.basis.mass_kg,
        basis_area_m2: m.basis.area_m2,
        basis_units: m.basis.units,
        basis_unit_desc: m.basis.unit_desc,
        basis_unit_weight_kg: m.basis.unit_weight_kg,
        facility_total_qty_std: m.production_std?.facility_total_qty_std ?? null,
        facility_total_unit_std: m.production_std?.facility_total_unit_std ?? null,
        product_direct_qty_std: m.production_std?.product_direct_qty_std ?? null,
        product_direct_unit_std: m.production_std?.product_direct_unit_std ?? null,
        period_product_output: m.production?.period_product_output ?? null
      }];
      addSheet(
        wb,
        '_meta_summary',
        metaRow,
        ["company","facility","period","product","sku","basis_type","basis_mass_kg","basis_area_m2","basis_units","basis_unit_desc","basis_unit_weight_kg","facility_total_qty_std","facility_total_unit_std","product_direct_qty_std","product_direct_unit_std","period_product_output"],
        { basis_mass_kg: "0.0000", basis_area_m2: "0.0000", basis_unit_weight_kg: "0.0000", facility_total_qty_std: "0.0000", product_direct_qty_std: "0.0000", period_product_output: "0.0000" }
      );

      return wb;
    }

    function addXlsxButton() {
      const btn = document.createElement('button');
      btn.id = 'downloadXlsx';
      btn.type = 'button';
      btn.textContent = 'Download XLSX';
      const zipBtn = document.querySelector('#downloadZip') || document.querySelector('#downloadJson');
      const container = zipBtn ? zipBtn.parentElement : document.body;
      container.appendChild(btn);

      btn.addEventListener('click', () => {
        try {
          const payload = buildExportJSON();
          const wb = buildWorkbookFromPayload(payload);
          window.XLSX.writeFile(wb, 'intake_export.xlsx');
          if (okMsg) {
            okMsg.textContent = 'XLSX downloaded.';
            setTimeout(() => okMsg.textContent = '', 2500);
          }
        } catch (err) {
          console.error(err);
          if (errors) {
            errors.textContent = 'Failed to build XLSX: ' + (err?.message || err);
          }
        }
      });
    }

    ensureXLSX(addXlsxButton);
  })();

  console.log('Initialization complete!');
}
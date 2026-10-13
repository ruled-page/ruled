import { describe, expect, it } from 'vitest';
import { readPivot, withSetup } from '../src/core/pivotTable';
import { rewritePivotCache, rewritePivotTable } from '../src/core/pivotWriter';
import { parseXml } from '../src/core/xml';

const CACHE =
  '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="1"><cacheSource type="worksheet"><worksheetSource ref="A1:D3" sheet="S"/></cacheSource><cacheFields count="4">' +
  '<cacheField name="Day" numFmtId="14"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="2026-10-01T00:00:00" maxDate="2026-10-01T00:00:00" count="1"><d v="2026-10-01T00:00:00"/></sharedItems></cacheField>' +
  '<cacheField name="What"><sharedItems/></cacheField>' +
  '<cacheField name="Amount"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1" maxValue="1"/></cacheField>' +
  '<cacheField name="Extra"><sharedItems/></cacheField>' +
  '</cacheFields></pivotCacheDefinition>';

describe('a pivot cache written again', () => {
  // 2 October 2026 is serial 46297.
  const records = [
    [46296, 'Tea', 1.5, true],
    [46297, null, 2, { error: '#N/A' }],
    [46297, 7, 3, null],
  ];

  it('writes each kind of value as a cache does, dates as dates, and says what kinds each field holds', () => {
    const cache = rewritePivotCache(CACHE, true, records, new Set([2]));
    expect(cache.definition).toContain(
      '<sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="2026-10-01T00:00:00" maxDate="2026-10-02T00:00:00" count="2"><d v="2026-10-01T00:00:00"/><d v="2026-10-02T00:00:00"/></sharedItems>',
    );
    expect(cache.definition).toContain('<sharedItems containsBlank="1" containsMixedTypes="1" containsNumber="1" containsInteger="1" minValue="7" maxValue="7" count="3"><s v="Tea"/><m/><n v="7"/></sharedItems>');
    expect(cache.definition).toContain('<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1.5" maxValue="3" count="3"><n v="1.5"/><n v="2"/><n v="3"/></sharedItems>');
    expect(cache.definition).toContain('<sharedItems containsBlank="1" count="3"><b v="1"/><e v="#N/A"/><m/></sharedItems>');
    expect(cache.records).toContain('<r><x v="1"/><x v="1"/><x v="1"/><x v="1"/></r>');
    expect([cache.index(0, 46297), cache.index(1, 'TEA'), cache.index(1, 'none'), cache.index(9, 1), cache.count(3), cache.count(9)]).toEqual([1, 0, -1, -1, 3, 0]);
  });

  it('keeps numbers out of the shared items unless listed, and has no records where it had none', () => {
    const cache = rewritePivotCache(CACHE, false, records, new Set());
    expect(cache.definition).toContain('<cacheField name="Amount"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1.5" maxValue="3"/></cacheField>');
    expect(cache.records).toBeUndefined();
    expect(cache.definition).not.toContain('recordCount="3"');
  });
});

describe('a pivot definition written again', () => {
  it('writes the Values field down the rows and leaves fields beyond the cache as they are', () => {
    const table = parseXml(
      '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" dataPosition="0"><location ref="A1:B2" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="5"><pivotField showAll="0"/><pivotField showAll="0" defaultSubtotal="0"/><pivotField showAll="0"/><pivotField showAll="0"/><pivotField/></pivotFields></pivotTableDefinition>',
    );
    const model = readPivot(table, parseXml(CACHE));
    const records = [
      [46296, 'Tea', 1.5, true],
      [46297, 'Cake', 2, false],
    ];
    const next = withSetup(model, { rows: [1, -2], cols: [], pages: [], data: [{ field: 2, fn: 'sum', name: 'Sum of Amount' }, { field: 2, fn: 'count', name: 'Count of Amount' }], hidden: {} }, records);
    const cache = rewritePivotCache(CACHE, false, records, new Set([1]));
    const xml = rewritePivotTable(
      '<pivotTableDefinition name="P" dataPosition="0"><location ref="A1:B2"/><pivotFields count="5"><pivotField showAll="0"/><pivotField showAll="0" defaultSubtotal="0"/><pivotField showAll="0"/><pivotField showAll="0"/><pivotField/></pivotFields></pivotTableDefinition>',
      { ...next, fields: next.fields.slice(0, 4) },
      { ref: 'A1:B7', firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 1 },
      cache,
    );
    expect(xml).toContain('<pivotTableDefinition name="P" dataOnRows="1">');
    // Cake sorts before Tea; no subtotal item where the field has none.
    expect(xml).toContain('<pivotField showAll="0" defaultSubtotal="0" axis="axisRow"><items count="2"><item x="1"/><item x="0"/></items></pivotField>');
    expect(xml).toContain('<pivotField/></pivotFields><rowFields count="2"><field x="1"/><field x="-2"/></rowFields>');
  });
});

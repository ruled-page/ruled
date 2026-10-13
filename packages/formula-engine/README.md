# @ruled/formula

Excel's calculation on its own, in a browser, a worker or Node. Give it
sheets of values and formulas, and it works them out as Excel does and keeps
the results up to date as cells change. It is the engine Ruled's editor,
`@ruled/sheet`, calculates with.

```ts
import { createWorkbook } from '@ruled/formula';

const book = createWorkbook({
  sheets: { Prices: [['Tea', 4.5], ['Cake', 3], ['Total', '=SUM(B1:B2)']] },
});

book.get('Prices!B3'); // 7.5
book.set('Prices!B1', 10); // [{ cell: 'Prices!B1', value: 10 }, { cell: 'Prices!B3', value: 13 }]
```

The documentation is at [ruled.page](https://ruled.page/formula/introduction/).

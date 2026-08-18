# Example style guide for a problem-and-solution handbook

- Preserve the source's exact wording, numeric values, notation, ordering, and typographical errors.
- Omit running headers, footers, and printed page numbers.
- Begin each problem with `\Prob{number}{title}`.
- Begin its matching solution with `\Sol{number}{title}`.
- Use plain nested `enumerate` environments; never hand-type `(a)`, `i.`, or `A.` labels.
- Use `align*` for derivation chains and `cases` for piecewise functions or distributions.
- Use the `Rcode` environment for R commands and console output.
- Use `booktabs` rules for tables.
- Reconstruct a graph with TikZ/pgfplots only when every plotted value, function, or summary statistic is visible in the page or extracted text.
- For boxplots, use `boxplot prepared` with the visible five-number summary; never estimate quartiles from pixels.
- Preserve source colours when they can be sampled or named reliably; otherwise flag the graph for review.
- Use `\ensuremath` inside reusable mathematical commands that may appear in table cells or ordinary text.

# Shared rich-text tab stops

Rich body and list paragraphs retain their original tab characters and emit
DrawingML tab positions from shared core rich fragments. Supply matching
measurement options to layout, preview and export. Tabs use four measured
spaces at the current run's size, relative to each accepted line. PptxGenJS
requires the tab options on individual rich runs as well as the containing
text shape; the converter supplies them without rewriting source runs.

`node test/rich-tabs.mjs` checks 12 body/list paragraphs across estimated,
measured and prepared-painting modes, compares emitted positions to shared
geometry, retains leading/trailing whitespace and checks native reimport still
contains tab characters. It does not prove native Office placement or full
rich-source/metadata round trips. Rich table-cell tab placement remains open.

The implementation is an unpublished draft stacked on shared furniture. It
requires coordinated core rich-tab fragments and the matching renderer; an
older npm dependency graph does not establish acceptance of these APIs.

---
type: added
---
RR-91: add `toPptx({slideNumberWithText: 'static'})` for a footer or header text with words around `{{slide.number}}` (`{{slide.number}} / {{deck.slideCount}}`). Keynote replaces the whole paragraph that holds a slide-number field, in the Slide Number placeholder or in a plain text box, with the number alone, so such a footer showed only `2`; `'static'` writes the number as fixed text in an ordinary tagged shape and Keynote shows `2 / 6`. The default `'field'` is unchanged (live field in the `sldNum` placeholder), a text that is only the number keeps its field in both modes, the authored text returns on import, and the total stays static because PowerPoint has no slide-count field.

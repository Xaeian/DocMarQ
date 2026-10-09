# Plan: silnik markdown, dokumenty i pisma w docmarq

Rozdzielamy `docmarq.md` tak samo jak pdfmarq w `0.6.0`: na czysty silnik markdown i dwa pakiety dokumentów.
`docmarq.paper` przejmuje dokumenty z banerem z `0.3.7`, a nowy `docmarq.legal` składa pisma, oświadczenia i umowy, z kolumnami i trybem równoległym.

Punktem wyjścia jest `0.3.7`, a wynikiem jedno wydanie `0.4.0`.
Pola stylów, sygnatury funkcji i nazwy modułów są te same co w pdfmarq, więc ten sam skrypt i md2paper przechodzą z PDF na DOCX zmianą importu.
Kod piszemy zgodnie z `c:\Python\pystyle.md`, dokumentację zgodnie z `c:\Python\style-md.md`.

## 1. Zasady

Te same co w pdfmarq:
- biblioteka składa dokument, nie treść, i nie zgaduje języka ani rodzaju dokumentu;
- frontmatter mówi, czym dokument jest, a styl, jak wygląda;
- każda warstwa importuje tylko warstwy pod sobą;
- domyślne wartości dają wygląd z `0.3.7`;
- widoczny błąd jest lepszy niż cichy zły wynik.

Różnica wobec pdfmarq: stronę składa Word.
Łamanie stron, wyrównanie kolumn i dzielenie wyrazów to jego natywne cechy, więc docmarq nie ma pomiaru próbnym renderem ani własnego łamania.

## 2. Warstwy

```text
docmarq.paper         dokumentacja: baner, baner kompaktowy w nagłówku, numer strony, podpisy
docmarq.legal       pisma: nagłówek pisma, preambuła, kolumny, końcówka, tryb równoległy
docmarq.md          silnik: markdown → treść w bieżącym celu zapisu
docmarq             rdzeń: DOCX, cel zapisu, sekcje z kolumnami, pola, podpisy, numer strony
```

`paper` i `legal` nie importują się nawzajem, czego pilnuje test importów.

## 3. Struktura plików

```text
docmarq/
  core.py              + into(cell), columns(n, gap), hyphenation(lang), heading_style(), heading(space_before=, space_after=)
  inline.py            + RichSegment(small_caps=)
  fields.py            nowy: pola Worda, także złożone (`IF NUMPAGES > 1`)
  parts.py             nowy: signatures(), page_number(), sign_labels(), LANG_LABELS
  md/
    renderer.py        MarkdownRenderer: parse(), index(), render_tokens(), render(), frontmatter_file(); top_blocks()
    frontmatter.py     split_frontmatter(), peek_frontmatter(), metadata_from_frontmatter(), docx_metadata()
    style.py           MarkdownStyle bez pól banera, podpisów i stopki; opcje z 5
    presets.py         tylko etykiety calloutów
  paper/               PaperStyle, md_to_docx(), lang_style(), baner, _skip_matching_h1()
  legal/               LegalStyle, md_to_docx(), md_parallel_to_docx(), lang_style(), nagłówek pisma
```

## 4. Mapowanie na Word

- **Cel zapisu:** rdzeń dopisuje bloki do treści dokumentu albo do komórki tabeli (`with doc.into(cell, width)`).
  Szerokość komórki zastępuje wtedy szerokość treści, więc tabele i obrazy w komórce mieszczą się w niej.
- **Kolumny:** `columns(n, gap)` otwiera sekcję ciągłą z `w:cols`, a `columns(1)` ją zamyka.
  Word sam wyrównuje kolumny przed ciągłym podziałem sekcji, także na ostatniej stronie długiej umowy.
- **Justowanie:** `w:jc="both"` na akapitach treści i punktów list.
- **Dzielenie wyrazów:** `w:autoHyphenation` w ustawieniach i język tekstu (`w:lang`) z presetu.
  Nagłówki, komórki tabel i nagłówek pisma dostają `w:suppressAutoHyphens`.
- **Kapitaliki:** `w:smallCaps`, kapitaliki Worda.
- **Numer strony:** pola `PAGE` i `NUMPAGES` w stopce.
  W `legal` całość siedzi w polu `IF NUMPAGES > 1`, więc pismo na jedną stronę nie ma numeru.
- **Logo obok tytułu:** obraz pływający (`wp:anchor`) z opływaniem kwadratowym przy marginesie.
  Tytuł ma po obu stronach wcięcie na szerokość logo, więc zostaje na środku strony, a preambuła opływa logo sama.
- **Podpisy:** tabela bez ramek, linie jako dolna krawędź akapitu, kropkowane w `legal`.
  `sign_bottom` robi z końcówki tabelę pływającą przy dolnym marginesie (`w:tblpPr`).
- **Tryb równoległy:** tabela bez ramek, kolumna na wersję, wiersz na blok.
  Wiersz może się złamać między stronami: z `w:cantSplit` Word ucina wiersz wyższy niż strona.
  Wiersz z nagłówkiem trzyma się następnego przez `keepNext`, a `<!-- pagebreak -->` kończy tabelę i zaczyna następną na nowej stronie.
  Tekst każdej wersji ma jej język (`w:lang`), więc Word dzieli wyrazy według właściwych reguł.

## 5. Opcje silnika

Każda opcja domyślnie daje wygląd z `0.3.7`.

- `justify`, `hyphenate` jak w pdfmarq.
- `keep_together = False`, bo `0.3.7` zostawiało łamanie Wordowi.
  `True` trzyma akapit w całości (`keepLines`).
  To jedyna świadoma różnica domyślnej wartości wobec pdfmarq.
- `h1_size`…`h6_size` (`None` to rozmiar treści), `h1_underline`, `h2_underline`, `head_mode` ustawiają style Heading 1…6.
- `head_gap_top`, `head_gap_bot` (`None` to odstępy z `0.3.7` dla każdego poziomu).
- Lista numerowana drukuje znacznik autora, `1.` albo `1)`.
  To jedyna zamierzona zmiana wyglądu w `paper` poza dwiema poprawkami z punktu 7.

## 6. Kolejność prac i odbiór

0. Wzorzec: 8 dokumentów z `0.3.7` jako części XML, z drugim przebiegiem dowodzącym determinizmu.
1. Rdzeń bez zmiany wyglądu: cel zapisu, kolumny, pola, kapitaliki, dzielenie wyrazów.
   Odbiór: wzorzec identyczny.
2. Podział: silnik, `paper`, `parts.py`, presety.
   Odbiór: wzorzec identyczny.
3. Opcje silnika z 5.
   Odbiór: wzorzec identyczny przy domyślnym stylu.
4. `legal`, jeden dokument.
5. Tryb równoległy.
6. Testy struktury OOXML, dokumentacja, `changes.md`, testy cross-lib w pdfmarq, przegląd pystyle i style-md.

Na tej maszynie nie ma Worda ani LibreOffice.
Testy sprawdzają strukturę OOXML, a wygląd w Wordzie ocenia użytkownik na przykładowych plikach.

## 7. Stan

Kroki 0–6 wykonane 2026-10-08, bez commita.
Wzorzec `0.3.7` zostaje identyczny w kanonicznym XML, z jedną przyjętą zmianą: `w:tblBorders` w kolejności schematu.

Decyzje z wdrożenia:

- `legal` włącza w stylu `Normal` kontrolę wdów i sierot, zeruje odstęp nad akapitem i składa tekst w `body_color`, jak pdfmarq.
- Preambuła kończy się akapitem 1pt z łamaniem `clear="all"`, więc klauzule zaczynają się pod pływającym logo.
- `head_mode` bez `Italic` zostawia pochylenie, które Word daje poziomom 4 i 6, bo tak wyglądało `0.3.7`.
- Poprawki: obraz nad linią `---` był gubiony jako pusty nagłówek, a ramki tabel bez ramek stały poza kolejnością schematu.
- Opisy z dawnego readme, których kod nie spełniał (biała lista HTML, tabele bez nagłówka, emoji, klikalne przypisy), poprawione w dokumentacji.
- W pdfmarq: test cross-lib na nowym API, zdanie o numeracji Worda w `md/guide.md`, nieaktualne przykłady w readme.

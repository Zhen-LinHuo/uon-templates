# latexmk configuration — University of Nottingham report template
# latexmk reads only `.latexmkrc` / `latexmkrc`, never a file named `.latex`.
$pdf_mode = 4;                          # 4 = lualatex (our default engine)
$bibtex_use = 2;
# biber resolves relative .bib paths against the .bcf's directory (build/),
# so point it at the sources and at the output directory.
$biber = 'biber --input-directory=. --output-directory=build %O %S';
$lualatex = 'lualatex -interaction=nonstopmode -file-line-error %O %S';
$out_dir = 'build';
$aux_dir = 'build';
# Keep %O %S: latexmk injects -recorder and the file name there.

@default_files = ('report');

# The shared brand layer lives in ../common. Keep the directory layout, or
# copy uon-brand.sty next to report.tex if you move this template.
$ENV{'TEXINPUTS'} = '../common//:' . ($ENV{'TEXINPUTS'} // '');

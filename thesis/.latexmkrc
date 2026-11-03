# latexmk configuration — University of Nottingham thesis template
# latexmk reads only `.latexmkrc` / `latexmkrc`, never a file named `.latex`.
# 4 = lualatex, 5 = xelatex, 1 = pdflatex. LuaLaTeX is our default engine.
$pdf_mode = 4;
$bibtex_use = 2;                       # run biber automatically when needed
# biber resolves relative .bib paths against the .bcf's directory (build/),
# so tell it where the sources are and where the .bbl must land.
$biber = 'biber --input-directory=. --output-directory=build %O %S';
$lualatex = 'lualatex -interaction=nonstopmode -file-line-error %O %S';
$out_dir = 'build';
$aux_dir = 'build';
# Keep %O %S: latexmk injects -recorder and the file name there.

# This directory holds several .tex files (rq.tex, titlepage.tex, chapters/).
# Without this, bare `latexmk` guesses the main file wrong.
@default_files = ('thesis');

# The shared brand layer lives in ../common. If you copy this template
# elsewhere, either keep the directory layout or copy uon-brand.sty next
# to thesis.tex.
$ENV{'TEXINPUTS'} = '../common//:' . ($ENV{'TEXINPUTS'} // '');

# Rebuild the acronym list without external tools (glossaries noidx mode).
$makeindex = '';

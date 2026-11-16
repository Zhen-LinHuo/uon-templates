# latexmk configuration — CEE annual report (LuaLaTeX + biber)
$pdf_mode = 4;                          # 4 = lualatex
$bibtex_use = 2;
$biber = 'biber --input-directory=. --output-directory=build %O %S';
$lualatex = 'lualatex -interaction=nonstopmode -file-line-error %O %S';
$out_dir = 'build';
$aux_dir = 'build';

@default_files = ('annual-report');

# The shared brand layer lives in ../common.
$ENV{'TEXINPUTS'} = '../common//:' . ($ENV{'TEXINPUTS'} // '');

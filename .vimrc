set encoding=utf8
set ffs=unix,dos,mac

filetype plugin on
filetype indent on

set nu
set autoread
set autoindent
set backspace=indent,eol,start
set complete-=i
set smarttab

set nrformats-=octal

set ttimeout
set ttimeoutlen=100

set laststatus=2
set ruler
set showcmd
set wildmenu

syntax enable 
set background=dark
try
    colorscheme murphy
catch
endtry

set nobackup
set nowb
set noswapfile

syn match   myTodo   contained   "\<\(TODO\|FIXME\):"
hi def link myTodo Todo

set ai 
set si 
set wrap 

set noeb vb t_vb=   " Stop it!!!

nnoremap <C-J> <C-W><C-J>
nnoremap <C-K> <C-W><C-K>
nnoremap <C-L> <C-W><C-L>
nnoremap <C-H> <C-W><C-H>

set splitbelow
set splitright

highlight OverLength ctermbg=red ctermfg=white guibg=#592929
match OverLength /\%81v.\+/

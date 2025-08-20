{ pkgs }: {
    deps = [
        pkgs.python312
        pkgs.python312Packages.pip
        pkgs.python312Packages.pandas
        pkgs.python312Packages.numpy
        pkgs.python312Packages.openpyxl
        pkgs.python312Packages.requests
        pkgs.ghostscript
        pkgs.tkinter
    ];
}
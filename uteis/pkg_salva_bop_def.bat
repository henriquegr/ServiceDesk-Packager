pdm_extract wspcol > wspcol.userload
pdm_extract wsptbl > wsptbl.userload

move /Y wspcol.userload ..\UserLoad\wspcol.userload
move /Y wsptbl.userload ..\UserLoad\wsptbl.userload

@echo off
@cscript pkg_salva_bop_def.vbs


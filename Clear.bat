@echo off
title Clear
echo.

if exist *.~?? del *.~??
if exist *.dcu del *.dcu
if exist *.opt del *.opt
if exist *.dsm del *.dsm
if exist *.dsk del *.dsk
if exist *.cfg del *.cfg
if exist *.ddp del *.ddp
if exist *.dof del *.dof
if exist *.bdsproj del *.bdsproj
if exist *.identcache del *.identcache
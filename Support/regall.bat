@echo off
for %%P in (*.ocx *.dll) do regsvr32 %%P /s
' ==UserScript==
' @name         Silently Run Command or Batch File from Same Directory
' @description  Silently call a .cmd or .bat file located in the same directory using the Windows Script Host Run Method.
' @remarks		Replace 'drive' with drive location, e.g. 'C'.
' @remarks		Replace 'scriptfolder' with script folder name, e.g. '_Scripts'.
' @remarks		Replace 'scriptname.cmd' with script name, e.g. 'scriptname.cmd' 'scriptname.bat'.
' @remarks	    intWindowStyle = 0, Hides the window and activates another window.
' @version      1.0
' @author       Dennis Murphy
' @namespace    https://dennismurphy.biz/
' ==/UserScript==

Set objWshShell = CreateObject("Wscript.Shell")

objWshShell.Run "drive:\scriptfolder\scriptname.cmd", 0
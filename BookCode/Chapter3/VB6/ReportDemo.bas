Attribute VB_Name = "Module1"
Option Explicit

Public objReport As Report

Public Const PRINTER_VIEW = 0
Public Const SCREEN_VIEW = 1

Public Const BOOL_TRUE = -1
Public Const BOOL_FALSE = 0
Public Const BOOL_NEITHER = 1

Public oConn As New ADODB.Connection


VERSION 5.00
Begin VB.UserControl FGraph3d 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "FGraph3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'FGraph 3d: pretty similiar to the 2d Fgraph, except, well, it's 3-d
'.this means it needs a different collection of Functions, as they will not have line colors but rather face colour properties and so forth

'also, the question of OpenGL or Direct3d comes to bear,
'as well as wether I should write a simple face display engine....
'in either case, we will need to build a list of vertexes and triangles.



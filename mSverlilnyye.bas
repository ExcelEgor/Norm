Attribute VB_Name = "mSverlilnyye"
Option Explicit

Function ZenkovaniyeFasokNaSverlilnovStanke(Diametr As Double, Optional VyderzhivaniyeRazmera As Boolean = True) As Double
    'МУНВ на работы выполняемые на сверлильных станках. Москва 2003
    'Карта 26. Зенкование фасок. Сталь конструкционная углеродистая

    ZenkovaniyeFasokNaSverlilnovStanke = IIf(Diametr > 20, 0.3, 0.25)

End Function

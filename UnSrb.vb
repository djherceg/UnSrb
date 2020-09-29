'
' UnSrb 1.0 // 29.9.2020 // Djordje Herceg
' Uklanja slova čćđšž iz stringa i zamenjuje ih slovima ccdjsz
'

Public Function UnSrb(s As String) As String
  Static ch As String, chm As String, cy As String, cym As String, dj As String, djm As String, sh As String, shm As String, zh As String, zhm As String
  
  ch = ChrW(&H10C)
  chm = ChrW(&H10D)
  cy = ChrW(&H106)
  cym = ChrW(&H107)
  dj = ChrW(&H110)
  djm = ChrW(&H111)
  sh = ChrW(&H160)
  shm = ChrW(&H161)
  zh = ChrW(&H17D)
  zhm = ChrW(&H17E)

  s = Replace(s, ch, "C")
  s = Replace(s, chm, "c")
  s = Replace(s, cy, "C")
  s = Replace(s, cym, "c")
  s = Replace(s, dj, "Dj")
  s = Replace(s, djm, "dj")
  s = Replace(s, sh, "S")
  s = Replace(s, shm, "s")
  s = Replace(s, zh, "Z")
  s = Replace(s, zhm, "z")
  s = Replace(s, " ", ".")
  
  UnSrb = s
End Function
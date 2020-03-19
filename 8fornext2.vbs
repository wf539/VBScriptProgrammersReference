Dim Index
Dim WordLength

WordLength = Len("elephant")

For Index = 1 to WordLength
  MsgBox Mid("elephant", Index, 1)
Next

MsgBox "elephant"
# PStrings

A string utility library for VBA (Visual Basic for Applications).

## Overview

All functions are prefixed with `pstr_` and live in the `pstr_PStrings` module.

## API Reference

### `pstr_FormatBytes(Amount)`

Formats a byte count into a human-readable string (KB, MB, GB).

```vb
Debug.Print pstr_FormatBytes(1023)       ' 1023 bytes
Debug.Print pstr_FormatBytes(1024)       ' 1.00 KB
Debug.Print pstr_FormatBytes(1048576)    ' 1.00 MB
Debug.Print pstr_FormatBytes(1073741824) ' 1.00 GB
```

---

### `pstr_Wrap(Text, Wrapper)`

Wraps a string with the given wrapper on both sides.

```vb
Debug.Print pstr_Wrap("ABC", Chr(34)) ' "ABC"
Debug.Print pstr_Wrap("ABC", "ABC")   ' ABCABCABC
```

---

### `pstr_IndexOf(Text, Char)`

Returns the index of the **first** occurrence of a character. Returns `-1` if not found.

```vb
Debug.Print pstr_IndexOf("hello", "e") ' 2
Debug.Print pstr_IndexOf("hello", "z") ' -1
```

---

### `pstr_LastIndexOf(Text, Char)`

Returns the index of the **last** occurrence of a character. Returns `-1` if not found.

```vb
Debug.Print pstr_LastIndexOf("hello", "l") ' 4
Debug.Print pstr_LastIndexOf("hello", "z") ' -1
```

---

### `pstr_CharAt(Text, [Index])`

Returns the character at the given index (0-based). Returns empty string if out of bounds.

```vb
Debug.Print pstr_CharAt("Brave new world", 0) ' B
Debug.Print pstr_CharAt("Brave new world", 4) ' e
Debug.Print pstr_CharAt("Brave new world", 999) ' (empty)
```

---

### `pstr_CharCodeAt(Text, [Index])`

Returns the Unicode (ASCII) value of the character at the given index. Returns `-1` for invalid input.

```vb
Debug.Print pstr_CharCodeAt("ABC", 0) ' 65
```

---

### `pstr_Slice(Text, StartIndex, [EndIndex])`

Returns a substring between `StartIndex` and `EndIndex` (exclusive). Supports negative indices.

```vb
Dim s As String: s = "The morning is upon us."
Debug.Print pstr_Slice(s, 1, 8)   ' "he morn"
Debug.Print pstr_Slice(s, 4, -2)  ' "morning is upon u"
Debug.Print pstr_Slice(s, 12)     ' "is upon us."
Debug.Print pstr_Slice(s, -3)     ' "us."
```

---

### `pstr_StartsWith(Text, Expression, [Compare])`

Returns `True` if `Text` starts with `Expression`.

```vb
Debug.Print pstr_StartsWith("Check", "che", vbTextCompare) ' True
Debug.Print pstr_StartsWith("Check", "che")                ' False
```

---

### `pstr_EndsWith(Text, Expression, [Compare])`

Returns `True` if `Text` ends with `Expression`.

```vb
Debug.Print pstr_EndsWith("Check", "ECK", vbTextCompare) ' True
Debug.Print pstr_EndsWith("Check", "ECK")                ' False
```

---

### `pstr_Join(Delimiter, Values...)`

Joins multiple strings with a delimiter.

```vb
Debug.Print pstr_Join(" ", "Joined string", "from two strings")
' Joined string from two strings
```

---

### `pstr_Trim(Text)`

Removes leading, trailing, and extra internal spaces (uses `WorksheetFunction.Trim`).

```vb
Debug.Print pstr_Trim("  String  with     whitespaces    ")
' >String with whitespaces<
```

---

### `pstr_JoinNonEmpty(Data, [Delimiter])`

Joins array elements, skipping empty values (`""` or `Empty`).

```vb
Dim Data As Variant: Data = Array("Value1", Empty, "Value2", "")
Debug.Print pstr_JoinNonEmpty(Data, ", ") ' Value1, Value2
```

---

### `pstr_FString(Text, Values...)`

Replaces `{0}`, `{1}`, ... placeholders with the given values. Also supports escape tokens:

| Token | Result |
|-------|--------|
| `{CrLf}` / `\\n` | `vbCrLf` / `vbNewLine` |
| `{Cr}` / `\\r` | `vbCr` |
| `{Lf}` | `vbLf` |
| `\\t` | `vbTab` |

```vb
Debug.Print pstr_FString("Hello, {0}! You are {1} years old.", "Alice", 30)
' Hello, Alice! You are 30 years old.
```

---

### `pstr_FormatString(Text, Values...)`

Replaces typed placeholders with values:

| Placeholder | Type |
|-------------|------|
| `%s` | String |
| `%d` | Numeric |
| `%t` | Date |

```vb
Debug.Print pstr_FormatString("Function: %s, Count: %d", "pstr_FormatString", 42)
' Function: pstr_FormatString, Count: 42
```

---

### `pstr_InString(Text, Compare, Values...)`

Returns `True` if any of `Values` is found in `Text`.

```vb
Debug.Print pstr_InString("Hello World", vbTextCompare, "world") ' True
```

---

### `pstr_IsEqual(Text1, Text2, [Compare])`

Returns `True` if two strings are equal. Defaults to `vbTextCompare`.

```vb
Debug.Print pstr_IsEqual("Hello", "hello") ' True
```

---

### `pstr_IsNullString(Expression)`

Returns `True` if the string is empty (`""` or `Empty`).

```vb
Debug.Print pstr_IsNullString("") ' True
```

## License

MIT

# Chart Property Structure & Defaults

## Property Hierarchy

```
ExtendedChartOptions
├── Position (InsertPosition) = PositionEnd
├── Anchor (string) = ""
├── ChartKind (ChartKind) = ChartKindColumn
│
├── Title (string) = ""
├── TitleOverlay (bool) = false
│
├── Categories ([]string) ✓ MANDATORY
│
├── Series ([]SeriesOptions) ✓ MANDATORY
│   ├── [0]
│   │   ├── Name (string) ✓ MANDATORY
│   │   ├── Values ([]float64) ✓ MANDATORY
│   │   ├── Color (string) = auto
│   │   ├── InvertIfNegative (bool) = false
│   │   ├── Smooth (bool) = false
│   │   ├── ShowMarkers (bool) = false
│   │   └── DataLabels (*DataLabelOptions) = nil
│   └── [1...n]
│
├── CategoryAxis (*AxisOptions) = defaults
│   ├── Title (string) = ""
│   ├── TitleOverlay (bool) = false
│   ├── Min (*float64) = nil (auto)
│   ├── Max (*float64) = nil (auto)
│   ├── MajorUnit (*float64) = nil (auto)
│   ├── MinorUnit (*float64) = nil (auto)
│   ├── Visible (bool) = true
│   ├── Position (AxisPosition) = AxisPositionBottom
│   ├── MajorTickMark (TickMark) = TickMarkOut
│   ├── MinorTickMark (TickMark) = TickMarkNone
│   ├── TickLabelPos (TickLabelPosition) = TickLabelNextTo
│   ├── NumberFormat (string) = "General"
│   ├── MajorGridlines (bool) = false
│   ├── MinorGridlines (bool) = false
│   └── CrossesAt (*float64) = nil (auto)
│
├── ValueAxis (*AxisOptions) = defaults (same structure as CategoryAxis)
│   └── MajorGridlines (bool) = true  ← Different default
│
├── Legend (*LegendOptions) = defaults
│   ├── Show (bool) = true
│   ├── Position (string) = "r"
│   └── Overlay (bool) = false
│
├── DataLabels (*DataLabelOptions) = defaults
│   ├── ShowValue (bool) = false
│   ├── ShowCategoryName (bool) = false
│   ├── ShowSeriesName (bool) = false
│   ├── ShowPercent (bool) = false
│   ├── ShowLegendKey (bool) = false
│   ├── Position (DataLabelPosition) = DataLabelBestFit
│   └── ShowLeaderLines (bool) = true
│
├── Properties (*ChartProperties) = defaults
│   ├── Style (ChartStyle) = 2
│   ├── RoundedCorners (bool) = false
│   ├── Date1904 (bool) = false
│   ├── Language (string) = "en-US"
│   ├── PlotVisibleOnly (bool) = true
│   ├── DisplayBlanksAs (string) = "gap"
│   └── ShowDataLabelsOverMax (bool) = false
│
├── BarChartOptions (*BarChartOptions) = defaults
│   ├── Direction (BarDirection) = BarDirectionColumn
│   ├── Grouping (BarGrouping) = BarGroupingClustered
│   ├── GapWidth (int) = 150
│   ├── Overlap (int) = 0
│   └── VaryColors (bool) = false
│
├── Width (int) = 6099523 EMUs (~6.5")
├── Height (int) = 3340467 EMUs (~3.5")
│
└── Caption (*CaptionOptions) = nil
```

---

## Property Value Ranges & Valid Options

### ChartKind (String representation for XML)
```
ChartKindColumn → XML: "barChart" with barDir="col"
ChartKindBar    → XML: "barChart" with barDir="bar"
ChartKindLine   → XML: "lineChart"
ChartKindPie    → XML: "pieChart"
ChartKindArea   → XML: "areaChart"
```

### ChartStyle (Integer 0-48)
```
Value Range: 0-48
  0  = No style
  1  = Style 1
  2  = Style 2 (DEFAULT - Professional)
  ...
  10 = Colorful
  ...
  42 = Monochrome
  ...
  48 = Style 48
```

### Position (enum InsertPosition)
```
PositionBeginning     Start of document
PositionEnd          End of document (DEFAULT)
PositionAfterText    After specified text (requires Anchor)
PositionBeforeText   Before specified text (requires Anchor)
```

### Legend Position (String)
```
"r"   Right (DEFAULT)
"l"   Left
"t"   Top
"b"   Bottom
"tr"  Top-right
```

### Axis Position (AxisPosition)
```
AxisPositionBottom  "b" - Bottom (DEFAULT for CategoryAxis)
AxisPositionLeft    "l" - Left (DEFAULT for ValueAxis)
AxisPositionRight   "r" - Right
AxisPositionTop     "t" - Top
```

### Tick Mark (TickMark)
```
TickMarkOut    "out"   Outside (DEFAULT for major)
TickMarkNone   "none"  None (DEFAULT for minor)
TickMarkIn     "in"    Inside
TickMarkCross  "cross" Cross
```

### Tick Label Position (TickLabelPosition)
```
TickLabelNextTo  "nextTo" Next to axis (DEFAULT)
TickLabelHigh    "high"   High
TickLabelLow     "low"    Low
TickLabelNone    "none"   No labels
```

### Bar Grouping (BarGrouping)
```
BarGroupingClustered       "clustered"      Side-by-side (DEFAULT)
BarGroupingStacked         "stacked"        Stacked
BarGroupingPercentStacked  "percentStacked" 100% stacked
BarGroupingStandard        "standard"       Standard
```

### Bar Direction (BarDirection)
```
BarDirectionColumn  "col" Vertical bars (DEFAULT)
BarDirectionBar     "bar" Horizontal bars
```

### Data Label Position (DataLabelPosition)
```
DataLabelBestFit     "bestFit" Auto positioning (DEFAULT)
DataLabelCenter      "ctr"     Center
DataLabelOutsideEnd  "outEnd"  Outside end
DataLabelInsideEnd   "inEnd"   Inside end
DataLabelInsideBase  "inBase"  Inside base
```

### Display Blanks As (String)
```
"gap"   Leave gap for missing data (DEFAULT)
"zero"  Treat blank as zero
"span"  Interpolate across gap
```

### Number Format Codes (String)
```
"General"    1234.5 (DEFAULT)
"0"          1235
"0.00"       1234.50
"#,##0"      1,235
"#,##0.00"   1,234.50
"$#,##0"     $1,235
"$#,##0.00"  $1,234.50
"0%"         50%
"0.0%"       50.5%
"0.00%"      50.45%
"0.00E+00"   1.23E+03
"[Red]0"     Negative in red
```

### Gap Width Range (Integer for BarChartOptions)
```
Value Range: 0-500
  0   = No gap
  150 = DEFAULT (150% gap)
  500 = Maximum gap
```

### Overlap Range (Integer for BarChartOptions)
```
Value Range: -100 to +100
  -100 = Maximum separation
  0    = DEFAULT (no overlap)
  +100 = Complete overlap
```

---

## Size Calculations (EMUs - English Metric Units)

### Conversion Formula
```
1 inch = 914,400 EMUs
1 cm = 360,000 EMUs

EMUs = inches * 914400
EMUs = cm * 360000
```

### Default Sizes
```
Width:  6,099,523 EMUs = 6.67 inches = 16.9 cm
Height: 3,340,467 EMUs = 3.65 inches = 9.3 cm
```

### Common Chart Sizes
```go
// Standard sizes
Width:  int(6.0 * 914400)  // 6" wide = 5,486,400 EMUs
Height: int(4.0 * 914400)  // 4" tall = 3,657,600 EMUs

// Full page width (US Letter - 8.5" with 1" margins)
Width:  int(6.5 * 914400)  // 6.5" = 5,943,600 EMUs

// Square
Width:  int(5.0 * 914400)  // 5" = 4,572,000 EMUs
Height: int(5.0 * 914400)  // 5" = 4,572,000 EMUs

// Widescreen (16:9 ratio)
Width:  int(6.4 * 914400)  // 6.4" = 5,852,160 EMUs
Height: int(3.6 * 914400)  // 3.6" = 3,291,840 EMUs
```

---

## Color Values (Hex RGB without #)

### Standard Office Theme Colors
```
Blue:       "4472C4"
Orange:     "ED7D31"
Gray:       "A5A5A5"
Yellow:     "FFC000"
Light Blue: "5B9BD5"
Green:      "70AD47"
Red:        "FF0000"
Purple:     "7030A0"
```

### How to Specify Colors
```go
Series: []SeriesOptions{
    {
        Name:   "Revenue",
        Values: []float64{100, 200, 150},
        Color:  "4472C4",  // Blue - NO # prefix
    },
}
```

---

## Validation Rules

### Categories
- ✓ Must have at least 1 category
- ✓ Can be any string (including dates, numbers as strings)
- ✓ All series must have values.length == categories.length

### Series
- ✓ Must have at least 1 series
- ✓ Each series must have a Name (non-empty string)
- ✓ Each series must have Values ([]float64)
- ✓ Values length must match Categories length

### Axes Numeric Values
```go
Min:       nil or *float64 (any value, typically >= 0)
Max:       nil or *float64 (must be > Min if both set)
MajorUnit: nil or *float64 (must be > 0 if set)
MinorUnit: nil or *float64 (must be > 0 and < MajorUnit if both set)
CrossesAt: nil or *float64 (any value)
```

### Size Constraints
```go
Width:  > 0 (typically 914400 to 9144000 = 1" to 10")
Height: > 0 (typically 914400 to 9144000 = 1" to 10")
```

---

## Example: Override Just One Default

```go
// Want everything default EXCEPT blue colored bars
updater.InsertChartExtended(ExtendedChartOptions{
    Categories: []string{"A", "B", "C"},
    Series: []SeriesOptions{
        {
            Name:   "Sales",
            Values: []float64{10, 20, 15},
            Color:  "4472C4",  // ← Only customization
        },
    },
    // All other properties use defaults
})
```

---

## Example: Common Customizations

### Professional Financial Chart
```go
updater.InsertChartExtended(ExtendedChartOptions{
    ChartKind: ChartKindColumn,
    Title:     "Quarterly Revenue",
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    Series: []SeriesOptions{
        {Name: "2025", Values: []float64{100000, 120000, 115000, 140000}},
        {Name: "2026", Values: []float64{110000, 135000, 125000, 160000}},
    },
    ValueAxis: &AxisOptions{
        Title:        "Revenue ($)",
        NumberFormat: "$#,##0",
    },
    Properties: &ChartProperties{
        Style: ChartStyle2,
    },
})
```

### Scientific Data with Grid
```go
minVal := 0.0
maxVal := 100.0

updater.InsertChartExtended(ExtendedChartOptions{
    ChartKind: ChartKindLine,
    Title:     "Temperature Over Time",
    Categories: []string{"0h", "6h", "12h", "18h", "24h"},
    Series: []SeriesOptions{
        {
            Name:        "Sensor A",
            Values:      []float64{20.5, 22.3, 25.7, 23.1, 21.8},
            Smooth:      true,
            ShowMarkers: true,
        },
    },
    ValueAxis: &AxisOptions{
        Title:          "Temperature (°C)",
        Min:            &minVal,
        Max:            &maxVal,
        NumberFormat:   "0.0",
        MajorGridlines: true,
        MinorGridlines: true,
    },
})
```

---

## Summary: What to Remember

1. **Only 4 properties are mandatory**: Categories, Series, Series.Name, Series.Values
2. **Everything has a default** - start simple, add customization as needed
3. **Nil pointers mean "use default"** - only create structs for properties you want to customize
4. **Units matter**: EMUs for size (914400 = 1"), hex for colors (no #), format codes for numbers
5. **Validation is automatic** - invalid values are caught with helpful error messages

---

## Quick Decision Tree

```
Do I need custom axes? 
  NO  → Use ChartOptions
  YES → Use ExtendedChartOptions with AxisOptions

Do I need data labels?
  NO  → Leave DataLabels nil
  YES → Set DataLabels.ShowValue = true

Do I need custom colors?
  NO  → Leave Series.Color empty (auto)
  YES → Set Series[i].Color = "4472C4"

Do I need stacked bars?
  NO  → Leave BarChartOptions nil
  YES → Set BarChartOptions.Grouping = BarGroupingStacked

Do I need specific axis range?
  NO  → Leave Min/Max nil (auto)
  YES → Set Min/Max to pointers: &minValue
```

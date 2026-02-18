# Chart Properties Quick Reference

## Mandatory Properties

Only these properties are **required** when creating a chart:

| Property | Type | Notes |
|----------|------|-------|
| `Categories` | `[]string` | ≥1 category required |
| `Series` | `[]SeriesData` / `[]SeriesOptions` | ≥1 series required |
| `Series[].Name` | `string` | Cannot be empty |
| `Series[].Values` | `[]float64` | Length must match categories |

**Everything else is optional with automatic defaults applied.**

---

## Default Values Summary

### Chart-Level Defaults

```go
// Applied automatically if not specified
ChartKind:          ChartKindColumn      // Column chart
Position:           PositionEnd          // Insert at end
Width:              6099523              // ~6.5 inches
Height:             3340467              // ~3.5 inches
```

### Chart Properties Defaults

```go
Properties: &ChartProperties{
    Style:               2,              // Professional style
    RoundedCorners:      false,
    Date1904:            false,          // Use 1900 date system
    Language:            "en-US",
    PlotVisibleOnly:     true,
    DisplayBlanksAs:     "gap",          // Gap for missing data
    ShowDataLabelsOverMax: false,
}
```

### Legend Defaults

```go
Legend: &LegendOptions{
    Show:     true,                     // Legend visible
    Position: "r",                      // Right side
    Overlay:  false,                    // Doesn't overlap chart
}
```

### Category Axis Defaults

```go
CategoryAxis: &AxisOptions{
    Visible:        true,
    Position:       AxisPositionBottom,  // Bottom
    MajorTickMark:  TickMarkOut,        // Outside ticks
    MinorTickMark:  TickMarkNone,
    TickLabelPos:   TickLabelNextTo,
    NumberFormat:   "General",
    MajorGridlines: false,               // No gridlines
    MinorGridlines: false,
    // Min, Max, Units: all auto-calculated
}
```

### Value Axis Defaults

```go
ValueAxis: &AxisOptions{
    Visible:        true,
    Position:       AxisPositionLeft,    // Left side
    MajorTickMark:  TickMarkOut,
    MinorTickMark:  TickMarkNone,
    TickLabelPos:   TickLabelNextTo,
    NumberFormat:   "General",
    MajorGridlines: true,                // Shows gridlines
    MinorGridlines: false,
    // Min, Max, Units: all auto-calculated
}
```

### Data Label Defaults

```go
DataLabels: &DataLabelOptions{
    ShowValue:        false,            // Hidden by default
    ShowCategoryName: false,
    ShowSeriesName:   false,
    ShowPercent:      false,
    ShowLegendKey:    false,
    Position:         "bestFit",
    ShowLeaderLines:  true,
}
```

### Bar/Column Chart Defaults

```go
BarChartOptions: &BarChartOptions{
    Direction:  BarDirectionColumn,     // "col" = vertical
    Grouping:   BarGroupingClustered,   // Side-by-side
    GapWidth:   150,                    // 150% gap
    Overlap:    0,                      // No overlap
    VaryColors: false,                  // Same color per series
}
```

### Series Defaults

```go
SeriesOptions{
    // Name, Values: REQUIRED
    Color:            "",                // Auto-assigned
    InvertIfNegative: false,
    Smooth:           false,             // (line charts)
    ShowMarkers:      false,             // (line charts)
    DataLabels:       nil,               // Uses global setting
}
```

---

## Common Chart Styles

| Style Value | Description |
|-------------|-------------|
| `0` | No style (minimal) |
| `1` | Style 1 (simple) |
| `2` | **Style 2 (default - professional)** |
| `10` | Colorful style |
| `42` | Monochrome |
| `1-48` | Full range of Office styles |

---

## Position Options

```go
PositionBeginning    // Start of document
PositionEnd          // End of document (default)
PositionAfterText    // After specific text (requires Anchor)
PositionBeforeText   // Before specific text (requires Anchor)
```

---

## Legend Positions

```go
"r"   // Right (default)
"l"   // Left
"t"   // Top
"b"   // Bottom
"tr"  // Top right (overlay)
```

---

## Axis Positions

```go
AxisPositionBottom   // "b" - Bottom (default for category)
AxisPositionLeft     // "l" - Left (default for value)
AxisPositionRight    // "r" - Right
AxisPositionTop      // "t" - Top
```

---

## Tick Mark Types

```go
TickMarkOut    // "out" - Outside (default for major)
TickMarkNone   // "none" - No marks (default for minor)
TickMarkIn     // "in" - Inside
TickMarkCross  // "cross" - Cross axis
```

---

## Bar Grouping Types

```go
BarGroupingClustered       // Side-by-side (default)
BarGroupingStacked         // Stacked
BarGroupingPercentStacked  // 100% stacked
```

---

## Data Label Positions

```go
DataLabelBestFit     // "bestFit" - Auto (default)
DataLabelCenter      // "ctr" - Center of data point
DataLabelOutsideEnd  // "outEnd" - Outside end
DataLabelInsideEnd   // "inEnd" - Inside end
DataLabelInsideBase  // "inBase" - Inside base
```

---

## Chart Types

```go
ChartKindColumn  // "barChart" with barDir="col" (default)
ChartKindBar     // "barChart" with barDir="bar"
ChartKindLine    // "lineChart"
ChartKindPie     // "pieChart"
ChartKindArea    // "areaChart"
```

---

## Number Format Examples

| Format | Result | Usage |
|--------|--------|-------|
| `"General"` | 1234.5 | Default |
| `"0"` | 1235 | Integers |
| `"0.00"` | 1234.50 | 2 decimals |
| `"#,##0"` | 1,235 | Thousands |
| `"$#,##0"` | $1,235 | Currency |
| `"0%"` | 50% | Percentage |
| `"0.0%"` | 50.5% | Percentage + decimal |

---

## Standard Office Colors (Hex)

```go
"4472C4"  // Blue
"ED7D31"  // Orange
"A5A5A5"  // Gray
"FFC000"  // Yellow
"5B9BD5"  // Light Blue
"70AD47"  // Green
"FF0000"  // Red
"7030A0"  // Purple
```

---

## Size Units (EMUs)

**English Metric Units (EMUs):**
- 1 inch = 914,400 EMUs
- Default width: 6,099,523 EMUs ≈ 6.67 inches
- Default height: 3,340,467 EMUs ≈ 3.65 inches

**Common sizes:**
```go
Width:  int(4.0 * 914400),   // 4 inches
Height: int(3.0 * 914400),   // 3 inches
```

---

## Minimal Example (All Defaults)

```go
updater.InsertChart(ChartOptions{
    Categories: []string{"A", "B", "C"},
    Series: []SeriesData{
        {Name: "Sales", Values: []float64{10, 20, 15}},
    },
})
```

**Result:** Column chart, right legend, no data labels, professional style, positioned at end.

---

## Full Customization Example

```go
minVal := 0.0
maxVal := 100.0

updater.InsertChartExtended(ExtendedChartOptions{
    Position:  PositionEnd,
    ChartKind: ChartKindColumn,
    Title:     "Sales Report",
    
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    
    Series: []SeriesOptions{
        {
            Name:   "Revenue",
            Values: []float64{80000, 95000, 87000, 102000},
            Color:  "4472C4",
        },
    },
    
    CategoryAxis: &AxisOptions{
        Title:        "Quarter",
    },
    
    ValueAxis: &AxisOptions{
        Title:          "Amount ($)",
        Min:            &minVal,
        Max:            &maxVal,
        NumberFormat:   "$#,##0",
        MajorGridlines: true,
    },
    
    Legend: &LegendOptions{
        Show:     true,
        Position: "r",
    },
    
    DataLabels: &DataLabelOptions{
        ShowValue: true,
        Position:  DataLabelOutsideEnd,
    },
    
    Properties: &ChartProperties{
        Style:          ChartStyle2,
        RoundedCorners: true,
        Language:       "en-US",
    },
    
    BarChartOptions: &BarChartOptions{
        Grouping: BarGroupingClustered,
        GapWidth: 100,
    },
})
```

---

## Key Takeaways

✅ **Only 4 properties are mandatory**: Categories, Series, Series.Name, Series.Values

✅ **All other properties have sensible defaults** that create professional-looking charts

✅ **Start simple**: Use basic `ChartOptions`, add customization only when needed

✅ **Defaults are Word-compatible**: Charts open correctly in MS Word and LibreOffice

✅ **Progressive enhancement**: Can always add more customization later without breaking existing code

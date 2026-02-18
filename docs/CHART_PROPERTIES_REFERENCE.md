## Chart Properties Reference

This document provides a comprehensive reference for all chart properties available in the DOCX Updater library, including mandatory fields, default values, and usage examples.

---

## Table of Contents

1. [Chart Creation Overview](#chart-creation-overview)
2. [Mandatory vs Optional Properties](#mandatory-vs-optional-properties)
3. [Property Categories](#property-categories)
4. [Default Values Reference](#default-values-reference)
5. [Extended Options](#extended-options)
6. [Usage Examples](#usage-examples)

---

## Chart Creation Overview

The library supports two levels of chart customization:

- **`ChartOptions`** - Simple, commonly-used properties (current implementation)
- **`ExtendedChartOptions`** - Comprehensive customization including axes, data labels, styles, etc. (new)

---

## Mandatory vs Optional Properties

### Mandatory Properties (Required)

| Property | Type | Description |
|----------|------|-------------|
| `Categories` | `[]string` | Category labels for X-axis. Must have at least 1 category. |
| `Series` | `[]SeriesData` or `[]SeriesOptions` | Data series. Must have at least 1 series. |
| `Series[].Name` | `string` | Name of each series. Cannot be empty. |
| `Series[].Values` | `[]float64` | Data values. Length must match Categories length. |

### Optional with Defaults

All other properties are optional and have sensible defaults applied automatically.

---

## Property Categories

### 1. Chart-Level Properties (`ChartProperties`)

Properties that apply to the entire chart.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Style` | `ChartStyle` | `2` | Chart style (1-48). Common values: 2 (professional), 10 (colorful), 42 (monochrome) |
| `RoundedCorners` | `bool` | `false` | Use rounded corners for chart area |
| `Date1904` | `bool` | `false` | Use 1904 date system (for Mac compatibility) |
| `Language` | `string` | `"en-US"` | Language code (e.g., "en-GB", "fr-FR") |
| `PlotVisibleOnly` | `bool` | `true` | Plot only visible cells from Excel data |
| `DisplayBlanksAs` | `string` | `"gap"` | How to display blanks: "gap", "zero", or "span" |
| `ShowDataLabelsOverMax` | `bool` | `false` | Show data labels even if over axis maximum |

**Example:**
```go
Properties: &ChartProperties{
    Style:          ChartStyle2,
    RoundedCorners: true,
    Language:       "en-GB",
}
```

---

### 2. Chart Positioning & Size

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Position` | `InsertPosition` | `PositionEnd` | Where to insert: PositionBeginning, PositionEnd, PositionAfterText, etc. |
| `Anchor` | `string` | `""` | Text to anchor relative positioning (for PositionAfterText/BeforeText) |
| `Width` | `int` | `6099523` | Width in EMUs (~6.5 inches). 914400 EMUs = 1 inch |
| `Height` | `int` | `3340467` | Height in EMUs (~3.5 inches) |

---

### 3. Chart Type & Appearance

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `ChartKind` | `ChartKind` | `ChartKindColumn` | Chart type: Column, Bar, Line, Pie, Area |
| `Title` | `string` | `""` | Main chart title (optional) |
| `TitleOverlay` | `bool` | `false` | Title overlays chart area

 (saves space) |

**Chart Types:**
- `ChartKindColumn` - Vertical bars (column chart)
- `ChartKindBar` - Horizontal bars (bar chart)
- `ChartKindLine` - Line chart
- `ChartKindPie` - Pie chart
- `ChartKindArea` - Area chart

---

### 4. Legend Options (`LegendOptions`)

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Show` | `bool` | `true` | Display legend |
| `Position` | `string` | `"r"` | Position: "r" (right), "l" (left), "t" (top), "b" (bottom), "tr" (top-right) |
| `Overlay` | `bool` | `false` | Legend overlays chart area |

**Example:**
```go
Legend: &LegendOptions{
    Show:     true,
    Position: "r",
    Overlay:  false,
}
```

---

### 5. Axis Options (`AxisOptions`)

Applies to both Category Axis (X) and Value Axis (Y).

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Title` | `string` | `""` | Axis title |
| `TitleOverlay` | `bool` | `false` | Title overlays chart |
| `Min` | `*float64` | `nil` (auto) | Minimum axis value |
| `Max` | `*float64` | `nil` (auto) | Maximum axis value |
| `MajorUnit` | `*float64` | `nil` (auto) | Major gridline interval |
| `MinorUnit` | `*float64` | `nil` (auto) | Minor gridline interval |
| `Visible` | `bool` | `true` | Show axis |
| `Position` | `AxisPosition` | (varies) | Axis position: "b" (bottom), "l" (left), "r" (right), "t" (top) |
| `MajorTickMark` | `TickMark` | `"out"` | Major tick style: "out", "in", "cross", "none" |
| `MinorTickMark` | `TickMark` | `"none"` | Minor tick style |
| `TickLabelPos` | `TickLabelPosition` | `"nextTo"` | Label position: "nextTo", "high", "low", "none" |
| `NumberFormat` | `string` | `"General"` | Number format code (e.g., "0.00", "#,##0") |
| `MajorGridlines` | `bool` | Value: `true`<br>Category: `false` | Show major gridlines |
| `MinorGridlines` | `bool` | `false` | Show minor gridlines |
| `CrossesAt` | `*float64` | `nil` (auto) | Where other axis crosses this one |

**Example:**
```go
ValueAxis: &AxisOptions{
    Title:          "Sales ($)",
    Min:            &minValue,  // minValue := 0.0
    Max:            &maxValue,  // maxValue := 10000.0
    MajorUnit:      &majorStep, // majorStep := 2000.0
    NumberFormat:   "#,##0",
    MajorGridlines: true,
    MinorGridlines: false,
}
```

---

### 6. Data Labels (`DataLabelOptions`)

Can be set globally or per-series.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `ShowValue` | `bool` | `false` | Show data point values |
| `ShowCategoryName` | `bool` | `false` | Show category names |
| `ShowSeriesName` | `bool` | `false` | Show series names |
| `ShowPercent` | `bool` | `false` | Show percentage (pie charts) |
| `ShowLegendKey` | `bool` | `false` | Show legend key symbol |
| `Position` | `DataLabelPosition` | `"bestFit"` | Position: "ctr", "inEnd", "inBase", "outEnd", "bestFit" |
| `ShowLeaderLines` | `bool` | `true` | Show leader lines (pie charts) |

**Example:**
```go
DataLabels: &DataLabelOptions{
    ShowValue:   true,
    Position:    DataLabelOutsideEnd,
    NumberFormat: "0.0",
}
```

---

### 7. Series Options (`SeriesOptions`)

Extended per-series customization.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Name` | `string` | **Required** | Series name |
| `Values` | `[]float64` | **Required** | Data values |
| `Color` | `string` | (auto) | Hex color code (e.g., "FF0000" for red) |
| `InvertIfNegative` | `bool` | `false` | Use different color for negative values |
| `Smooth` | `bool` | `false` | Smooth line curves (line charts) |
| `ShowMarkers` | `bool` | `false` | Show data point markers (line charts) |
| `DataLabels` | `*DataLabelOptions` | `nil` | Override global data labels for this series |

**Example:**
```go
Series: []SeriesOptions{
    {
        Name:             "Q1 Sales",
        Values:           []float64{100, 150, 120},
        Color:            "4472C4",
        InvertIfNegative: true,
    },
}
```

---

### 8. Bar/Column Chart Options (`BarChartOptions`)

Specific to bar and column charts.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Direction` | `BarDirection` | `"col"` | "col" (vertical), "bar" (horizontal) |
| `Grouping` | `BarGrouping` | `"clustered"` | "clustered", "stacked", "percentStacked" |
| `GapWidth` | `int` | `150` | Gap between bar groups (0-500%) |
| `Overlap` | `int` | `0` | Bar overlap (-100 to 100%) |
| `VaryColors` | `bool` | `false` | Use different color for each data point |

**Example:**
```go
BarChartOptions: &BarChartOptions{
    Direction:  BarDirectionColumn,
    Grouping:   BarGroupingStacked,
    GapWidth:   100,
    Overlap:    0,
    VaryColors: false,
}
```

---

## Default Values Reference

### Quick Reference Table

| Category | Property | Default Value |
|----------|----------|---------------|
| **Chart** | Style | `2` |
| **Chart** | Language | `"en-US"` |
| **Chart** | Date1904 | `false` |
| **Chart** | RoundedCorners | `false` |
| **Position** | Position | `PositionEnd` |
| **Size** | Width | `6099523 EMUs` (~6.5") |
| **Size** | Height | `3340467 EMUs` (~3.5") |
| **Legend** | Show | `true` |
| **Legend** | Position | `"r"` (right) |
| **Legend** | Overlay | `false` |
| **Axis** | Visible | `true` |
| **Axis** | MajorTickMark | `"out"` |
| **Axis** | MinorTickMark | `"none"` |
| **Axis** | TickLabelPos | `"nextTo"` |
| **Axis** | MajorGridlines (Value) | `true` |
| **Axis** | MajorGridlines (Category) | `false` |
| **Data Labels** | ShowValue | `false` |
| **Data Labels** | Position | `"bestFit"` |
| **Bar Chart** | Direction | `"col"` |
| **Bar Chart** | Grouping | `"clustered"` |
| **Bar Chart** | GapWidth | `150` |
| **Bar Chart** | Overlap | `0` |

---

## Extended Options

### Using ExtendedChartOptions

For full control over chart appearance:

```go
err := updater.InsertChartExtended(ExtendedChartOptions{
    Position:  PositionEnd,
    ChartKind: ChartKindColumn,
    Title:     "Sales Performance",
    
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    
    Series: []SeriesOptions{
        {
            Name:   "Revenue",
            Values: []float64{100000, 150000, 130000, 180000},
            Color:  "4472C4",
        },
        {
            Name:   "Profit",
            Values: []float64{20000, 35000, 28000, 45000},
            Color:  "70AD47",
        },
    },
    
    CategoryAxis: &AxisOptions{
        Title:        "Quarter",
        MajorTickMark: TickMarkOut,
    },
    
    ValueAxis: &AxisOptions{
        Title:          "Amount ($)",
        NumberFormat:   "$#,##0",
        MajorGridlines: true,
    },
    
    Legend: &LegendOptions{
        Show:     true,
        Position: "r",
    },
    
    Properties: &ChartProperties{
        Style:          ChartStyle2,
        RoundedCorners: true,
        Language:       "en-US",
    },
    
    BarChartOptions: &BarChartOptions{
        Direction: BarDirectionColumn,
        Grouping:  BarGroupingClustered,
        GapWidth:  150,
    },
})
```

---

## Usage Examples

### Example 1: Simple Chart (Current Implementation)

```go
err := updater.InsertChart(ChartOptions{
    Position:          PositionEnd,
    ChartKind:         ChartKindColumn,
    Title:             "Monthly Sales",
    CategoryAxisTitle: "Month",
    ValueAxisTitle:    "Sales ($)",
    Categories:        []string{"Jan", "Feb", "Mar"},
    Series: []SeriesData{
        {Name: "Revenue", Values: []float64{10000, 15000, 12000}},
    },
    ShowLegend:     true,
    LegendPosition: "r",
})
```

### Example 2: Stacked Bar Chart with Data Labels

```go
err := updater.InsertChartExtended(ExtendedChartOptions{
    ChartKind: ChartKindColumn,
    Title:     "Product Sales by Region",
    
    Categories: []string{"North", "South", "East", "West"},
    
    Series: []SeriesOptions{
        {Name: "Product A", Values: []float64{100, 120, 90, 110}, Color: "4472C4"},
        {Name: "Product B", Values: []float64{80, 95, 105, 90}, Color: "ED7D31"},
    },
    
    DataLabels: &DataLabelOptions{
        ShowValue: true,
        Position:  DataLabelCenter,
    },
    
    BarChartOptions: &BarChartOptions{
        Grouping: BarGroupingStacked,
    },
    
    ValueAxis: &AxisOptions{
        Title:        "Units Sold",
        NumberFormat: "#,##0",
    },
})
```

### Example 3: Line Chart with Customized Axes

```go
minVal := 0.0
maxVal := 100.0
majorUnit := 20.0

err := updater.InsertChartExtended(ExtendedChartOptions{
    ChartKind: ChartKindLine,
    Title:     "Temperature Trends",
    
    Categories: []string{"Mon", "Tue", "Wed", "Thu", "Fri"},
    
    Series: []SeriesOptions{
        {
            Name:        "High",
            Values:      []float64{72, 75, 78, 74, 76},
            Color:       "FF0000",
            Smooth:      true,
            ShowMarkers: true,
        },
        {
            Name:        "Low",
            Values:      []float64{58, 60, 62, 59, 61},
            Color:       "0000FF",
            Smooth:      true,
            ShowMarkers: true,
        },
    },
    
    ValueAxis: &AxisOptions{
        Title:          "Temperature (°F)",
        Min:            &minVal,
        Max:            &maxVal,
        MajorUnit:      &majorUnit,
        MajorGridlines: true,
        MinorGridlines: true,
        NumberFormat:   "0°",
    },
})
```

### Example 4: Pie Chart with Percentages

```go
err := updater.InsertChartExtended(ExtendedChartOptions{
    ChartKind: ChartKindPie,
    Title:     "Market Share",
    
    Categories: []string{"Company A", "Company B", "Company C", "Others"},
    
    Series: []SeriesOptions{
        {
            Name:   "Share",
            Values: []float64{35, 28, 22, 15},
        },
    },
    
    DataLabels: &DataLabelOptions{
        ShowPercent:     true,
        ShowCategoryName: true,
        Position:        DataLabelBestFit,
        ShowLeaderLines: true,
    },
    
    Legend: &LegendOptions{
        Show:     true,
        Position: "r",
    },
})
```

---

## Chart Style Gallery

Office provides 48 predefined chart styles. Common ones:

| Style | Description |
|-------|-------------|
| `1` | Simple, minimal styling |
| `2` | Professional (default in templates) |
| `3-9` | Various color schemes |
| `10` | Colorful, vibrant |
| `11-41` | Various professional styles |
| `42` | Monochrome |
| `43-48` | High-contrast styles |

---

## Color Reference

### Standard Office Colors (Hex Codes)

| Color | Hex Code |
|-------|----------|
| Blue | `4472C4` |
| Orange | `ED7D31` |
| Gray | `A5A5A5` |
| Yellow | `FFC000` |
| Light Blue | `5B9BD5` |
| Green | `70AD47` |
| Red | `FF0000` |
| Purple | `7030A0` |

---

## Number Format Codes

Common format codes for axes and data labels:

| Format Code | Display | Example |
|-------------|---------|---------|
| `General` | Default | 1234.5 |
| `0` | Integer | 1235 |
| `0.00` | 2 decimals | 1234.50 |
| `#,##0` | Thousands separator | 1,235 |
| `#,##0.00` | Thousands + decimals | 1,234.50 |
| `$#,##0` | Currency | $1,235 |
| `0%` | Percentage | 50% |
| `0.0%` | Percentage + decimal | 50.5% |
| `0.00E+00` | Scientific | 1.23E+03 |
| `m/d/yyyy` | Date | 2/18/2026 |

---

## Best Practices

1. **Keep it simple**: Start with basic `ChartOptions`, only use `ExtendedChartOptions` when needed
2. **Color consistency**: Use your brand colors consistently across series
3. **Axis ranges**: Set Min/Max when you need consistent scales across multiple charts
4. **Data labels**: Use sparingly - too many labels clutter the chart
5. **Legend**: Place on right for most charts, bottom for wide charts
6. **Gridlines**: Major gridlines help readability, minor gridlines can clutter
7. **Number formats**: Always format currency and percentages appropriately

---

## Migration from ChartOptions to ExtendedChartOptions

Current code:
```go
ChartOptions{
    Title:             "Sales",
    CategoryAxisTitle: "Month",
    ValueAxisTitle:    "Amount",
    ShowLegend:        true,
    LegendPosition:    "r",
}
```

Extended equivalent:
```go
ExtendedChartOptions{
    Title: "Sales",
    CategoryAxis: &AxisOptions{Title: "Month"},
    ValueAxis:    &AxisOptions{Title: "Amount"},
    Legend:       &LegendOptions{Show: true, Position: "r"},
}
```

---

## Conclusion

The extended chart options provide professional-grade chart customization while maintaining backward compatibility with the simple `ChartOptions` API. Use the level of customization appropriate for your needs.

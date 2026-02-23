package godocx

import (
	"strings"
	"testing"
)

func TestGenerateScatterChartXML(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindScatter,
		Categories: []string{"A", "B", "C"},
		Series: []SeriesOptions{
			{Name: "Series1", Values: []float64{10, 20, 30}},
		},
		ScatterChartOptions: &ScatterChartOptions{
			ScatterStyle: "smoothMarker",
			VaryColors:   true,
		},
	}

	result := generateScatterChartXML(opts)

	if !strings.Contains(result, "<c:scatterChart>") {
		t.Error("expected scatterChart element")
	}
	if !strings.Contains(result, `<c:scatterStyle val="smoothMarker"/>`) {
		t.Error("expected smoothMarker scatter style")
	}
	if !strings.Contains(result, `<c:varyColors val="1"/>`) {
		t.Error("expected varyColors=1")
	}
	if !strings.Contains(result, "</c:scatterChart>") {
		t.Error("expected closing scatterChart")
	}
	if !strings.Contains(result, "<c:axId") {
		t.Error("expected axis IDs")
	}
}

func TestGenerateScatterChartXML_DefaultStyle(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindScatter,
		Categories: []string{"A"},
		Series: []SeriesOptions{
			{Name: "S1", Values: []float64{5}},
		},
	}

	result := generateScatterChartXML(opts)

	if !strings.Contains(result, `<c:scatterStyle val="marker"/>`) {
		t.Error("expected default marker style")
	}
	if !strings.Contains(result, `<c:varyColors val="0"/>`) {
		t.Error("expected varyColors=0 by default")
	}
}

func TestGenerateScatterChartXML_WithDataLabels(t *testing.T) {
	opts := ChartOptions{
		ChartKind:  ChartKindScatter,
		Categories: []string{"X1"},
		Series: []SeriesOptions{
			{Name: "S1", Values: []float64{10}},
		},
		DataLabels: &DataLabelOptions{
			ShowValue: true,
			Position:  DataLabelBestFit,
		},
	}

	result := generateScatterChartXML(opts)

	if !strings.Contains(result, `<c:showVal val="1"/>`) {
		t.Error("expected data labels with showVal")
	}
}

func TestGenerateScatterSeriesXML(t *testing.T) {
	series := SeriesOptions{
		Name:   "TestSeries",
		Values: []float64{1.5, 2.5, 3.5},
		Color:  "FF0000",
	}
	opts := ChartOptions{
		Categories: []string{"P1", "P2", "P3"},
		Series:     []SeriesOptions{series},
	}

	result := generateScatterSeriesXML(0, series, opts)

	if !strings.Contains(result, "<c:ser>") {
		t.Error("expected ser element")
	}
	if !strings.Contains(result, "TestSeries") {
		t.Error("expected series name")
	}
	if !strings.Contains(result, `<c:xVal>`) {
		t.Error("expected xVal element")
	}
	if !strings.Contains(result, `<c:yVal>`) {
		t.Error("expected yVal element")
	}
	if !strings.Contains(result, "FF0000") {
		t.Error("expected color in output")
	}
}

func TestGenerateScatterSeriesXML_WithXValues(t *testing.T) {
	series := SeriesOptions{
		Name:    "Scatter",
		Values:  []float64{10, 20, 30},
		XValues: []float64{1.5, 3.0, 4.5},
	}
	opts := ChartOptions{
		Categories: []string{"A", "B", "C"},
		Series:     []SeriesOptions{series},
		ScatterChartOptions: &ScatterChartOptions{
			ScatterStyle: "line",
		},
	}

	result := generateScatterSeriesXML(0, series, opts)

	if !strings.Contains(result, "1.5") {
		t.Error("expected XValues in output")
	}
	if !strings.Contains(result, "4.5") {
		t.Error("expected XValues in output")
	}
}

func TestGenerateScatterSeriesXML_SmoothLine(t *testing.T) {
	series := SeriesOptions{
		Name:   "Smooth",
		Values: []float64{5, 10},
	}
	opts := ChartOptions{
		Categories: []string{"A", "B"},
		Series:     []SeriesOptions{series},
		ScatterChartOptions: &ScatterChartOptions{
			ScatterStyle: "smooth",
		},
	}

	result := generateScatterSeriesXML(0, series, opts)

	if !strings.Contains(result, `<c:smooth val="1"/>`) {
		t.Error("expected smooth=1 for smooth style")
	}
}

func TestGenerateScatterSeriesXML_WithPerSeriesLabels(t *testing.T) {
	series := SeriesOptions{
		Name:   "Labeled",
		Values: []float64{5},
		DataLabels: &DataLabelOptions{
			ShowValue: true,
		},
	}
	opts := ChartOptions{
		Categories: []string{"A"},
		Series:     []SeriesOptions{series},
	}

	result := generateScatterSeriesXML(0, series, opts)

	if !strings.Contains(result, `<c:showVal val="1"/>`) {
		t.Error("expected per-series data labels")
	}
}

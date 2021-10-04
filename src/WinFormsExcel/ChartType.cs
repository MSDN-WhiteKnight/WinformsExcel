/* WinForms Excel library 
 * Copyright (c) 2021,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Text;
/*Windows Forms Excel Library - ChartType enumeration
 http://smallsoft2.blogspot.ru/
 */

namespace ExtraControls
{
    /// <summary>
    /// Represents Excel chart type
    /// </summary>
    /// <remarks>
    /// This enumeration is a copy of XlChartType (https://docs.microsoft.com/en-us/office/vba/api/excel.xlcharttype) 
    /// from Excel API, provided to enable the charting support in WinFormsExcel libary.
    /// </remarks>
    public enum ChartType
    {
        /// <summary>
        /// Scatter
        /// </summary>
        xlXYScatter = -4169,

        /// <summary>
        /// Radar
        /// </summary>
        xlRadar = -4151,

        /// <summary>
        /// Doughnut
        /// </summary>
        xlDoughnut = -4120,

        /// <summary>
        /// 3D Pie
        /// </summary>
        xl3DPie = -4102,

        /// <summary>
        /// 3D Line
        /// </summary>
        xl3DLine = -4101,

        /// <summary>
        /// 3D Column
        /// </summary>
        xl3DColumn = -4100,

        /// <summary>
        /// 3D Area
        /// </summary>
        xl3DArea = -4098,

        /// <summary>
        /// Area
        /// </summary>
        xlArea = 1,

        /// <summary>
        /// Line
        /// </summary>
        xlLine = 4,

        /// <summary>
        /// Pie
        /// </summary>
        xlPie = 5,

        /// <summary>
        /// Bubble
        /// </summary>
        xlBubble = 15,

        /// <summary>
        /// Clustered Column
        /// </summary>
        xlColumnClustered = 51,

        /// <summary>
        /// Stacked Column
        /// </summary>
        xlColumnStacked = 52,

        /// <summary>
        /// 100% Stacked Column.
        /// </summary>
        xlColumnStacked100 = 53,

        /// <summary>
        /// Clustered Cone Bar
        /// </summary>
        xl3DColumnClustered = 54,

        /// <summary>
        /// Stacked Cone Bar
        /// </summary>
        xl3DColumnStacked = 55,

        /// <summary>
        /// 100% Stacked Cone Bar
        /// </summary>
        xl3DColumnStacked100 = 56,

        /// <summary>
        /// Clustered Bar
        /// </summary>
        xlBarClustered = 57,

        /// <summary>
        /// Stacked Bar
        /// </summary>
        xlBarStacked = 58,

        /// <summary>
        /// 100% Stacked Bar
        /// </summary>
        xlBarStacked100 = 59,

        /// <summary>
        /// 3D Clustered Bar
        /// </summary>
        xl3DBarClustered = 60,

        /// <summary>
        /// 3D Stacked Bar
        /// </summary>
        xl3DBarStacked = 61,

        /// <summary>
        /// 3D 100% Stacked Bar
        /// </summary>
        xl3DBarStacked100 = 62,

        /// <summary>
        /// Stacked Line
        /// </summary>
        xlLineStacked = 63,

        /// <summary>
        /// 100% Stacked Line
        /// </summary>
        xlLineStacked100 = 64,

        /// <summary>
        /// Line with Markers
        /// </summary>
        xlLineMarkers = 65,

        /// <summary>
        /// Stacked Line with Markers
        /// </summary>
        xlLineMarkersStacked = 66,

        /// <summary>
        /// 100% Stacked Line with Markers
        /// </summary>
        xlLineMarkersStacked100 = 67,

        /// <summary>
        /// Pie of Pie
        /// </summary>
        xlPieOfPie = 68,

        /// <summary>
        /// Exploded Pie
        /// </summary>
        xlPieExploded = 69,

        /// <summary>
        /// Exploded 3D Pie
        /// </summary>
        xl3DPieExploded = 70,

        /// <summary>
        /// Bar of Pie
        /// </summary>
        xlBarOfPie = 71,

        /// <summary>
        /// Scatter with Smoothed Lines
        /// </summary>
        xlXYScatterSmooth = 72,

        /// <summary>
        /// Scatter with Smoothed Lines and No Data Markers
        /// </summary>
        xlXYScatterSmoothNoMarkers = 73,

        /// <summary>
        /// Scatter with Lines
        /// </summary>
        xlXYScatterLines = 74,

        /// <summary>
        /// Scatter with Lines and No Data Markers
        /// </summary>
        xlXYScatterLinesNoMarkers = 75,

        /// <summary>
        /// Stacked Area
        /// </summary>
        xlAreaStacked = 76,

        /// <summary>
        /// 100% Stacked Area
        /// </summary>
        xlAreaStacked100 = 77,

        /// <summary>
        /// 3D Stacked Area
        /// </summary>
        xl3DAreaStacked = 78,

        /// <summary>
        /// 100% Stacked Area
        /// </summary>
        xl3DAreaStacked100 = 79,

        /// <summary>
        /// Exploded Doughnut
        /// </summary>
        xlDoughnutExploded = 80,

        /// <summary>
        /// Radar with Data Markers
        /// </summary>
        xlRadarMarkers = 81,

        /// <summary>
        /// Filled Radar
        /// </summary>
        xlRadarFilled = 82,

        /// <summary>
        /// 3D Surface
        /// </summary>
        xlSurface = 83,

        /// <summary>
        /// 3D Surface (wireframe)
        /// </summary>
        xlSurfaceWireframe = 84,

        /// <summary>
        /// Surface (Top View)
        /// </summary>
        xlSurfaceTopView = 85,

        /// <summary>
        /// Surface (Top View wireframe)
        /// </summary>
        xlSurfaceTopViewWireframe = 86,

        /// <summary>
        /// Bubble with 3D effects
        /// </summary>
        xlBubble3DEffect = 87,

        /// <summary>
        /// High-Low-Close
        /// </summary>
        xlStockHLC = 88,

        /// <summary>
        /// Open-High-Low-Close
        /// </summary>
        xlStockOHLC = 89,

        /// <summary>
        /// Volume-High-Low-Close
        /// </summary>
        xlStockVHLC = 90,

        /// <summary>
        /// Volume-Open-High-Low-Close
        /// </summary>
        xlStockVOHLC = 91,

        /// <summary>
        /// Clustered Cone Column
        /// </summary>
        xlCylinderColClustered = 92,

        /// <summary>
        /// Stacked Cone Column
        /// </summary>
        xlCylinderColStacked = 93,

        /// <summary>
        /// 100% Stacked Cylinder Column
        /// </summary>
        xlCylinderColStacked100 = 94,

        /// <summary>
        /// Clustered Cylinder Bar
        /// </summary>
        xlCylinderBarClustered = 95,

        /// <summary>
        /// Stacked Cylinder Bar
        /// </summary>
        xlCylinderBarStacked = 96,

        /// <summary>
        /// 100% Stacked Cylinder Bar
        /// </summary>
        xlCylinderBarStacked100 = 97,

        /// <summary>
        /// 3D Cylinder Column
        /// </summary>
        xlCylinderCol = 98,

        /// <summary>
        /// Clustered Cone Column
        /// </summary>
        xlConeColClustered = 99,

        /// <summary>
        /// Stacked Cone Column
        /// </summary>
        xlConeColStacked = 100,

        /// <summary>
        /// 100% Stacked Cone Column
        /// </summary>
        xlConeColStacked100 = 101,

        /// <summary>
        /// Clustered Cone Bar
        /// </summary>
        xlConeBarClustered = 102,

        /// <summary>
        /// Stacked Cone Bar
        /// </summary>
        xlConeBarStacked = 103,

        /// <summary>
        /// 100% Stacked Cone Bar
        /// </summary>
        xlConeBarStacked100 = 104,

        /// <summary>
        /// 3D Cone Column
        /// </summary>
        xlConeCol = 105,

        /// <summary>
        /// Clustered Pyramid Column
        /// </summary>
        xlPyramidColClustered = 106,

        /// <summary>
        /// Stacked Pyramid Column
        /// </summary>
        xlPyramidColStacked = 107,

        /// <summary>
        /// 100% Stacked Pyramid Column
        /// </summary>
        xlPyramidColStacked100 = 108,

        /// <summary>
        /// Clustered Pyramid Bar
        /// </summary>
        xlPyramidBarClustered = 109,

        /// <summary>
        /// Stacked Pyramid Bar
        /// </summary>
        xlPyramidBarStacked = 110,

        /// <summary>
        /// 100% Stacked Pyramid Bar
        /// </summary>
        xlPyramidBarStacked100 = 111,

        /// <summary>
        /// 3D Pyramid Column
        /// </summary>
        xlPyramidCol = 112,
    }
}

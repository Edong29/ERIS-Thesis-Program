Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Windows.Forms

Imports OxyPlot
Imports OxyPlot.Series
Imports OxyPlot.WindowsForms
Imports DataPoint = OxyPlot.DataPoint
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock
Imports OxyPlot.Axes
Imports System.Timers
'Imports System.Windows.Media

Public Class Form1
    Dim myPort As Array
    Dim serialPorts() As SerialPort = New SerialPort(2) {}
    Dim isStarted As Boolean = False ' Flag to track if "start" command has been sent
    Dim filePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\IJ's Project\Recordings\Continuous Data\continuousData.txt"
    'Dim triggeredFolderPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "IJ's Project\Recordings\Triggered Data")
    Dim triggeredFolderPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)) 'save triggered data to desktop by default.
    Dim triggeredFilePath As String = "" ' File path for triggered data recording
    Dim isTriggered As Boolean = False ' Flag to track if triggered recording is active
    Dim triggerStartTime As DateTime ' Time when trigger condition started
    Dim triggerDuration As TimeSpan = TimeSpan.FromSeconds(40) ' Duration for triggered recording (1 minute)

    Private timeSeries1 As LineSeries
    Private xSeries1 As LineSeries
    Private ySeries1 As LineSeries
    Private zSeries1 As LineSeries
    Private timeSeries2 As LineSeries
    Private xSeries2 As LineSeries
    Private ySeries2 As LineSeries
    Private zSeries2 As LineSeries
    Private timeSeries3 As LineSeries
    Private xSeries3 As LineSeries
    Private ySeries3 As LineSeries
    Private zSeries3 As LineSeries

    ' This method handles the form load event
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate COM ports
        myPort = SerialPort.GetPortNames()
        RoofDeckComPortComboBox.Items.AddRange(myPort)
        MidHeightComPortComboBox.Items.AddRange(myPort)
        GroundFloorComPortComboBox.Items.AddRange(myPort)

        ' Populate baud rates
        'Dim baudRates() As Integer = {9600, 14400, 19200, 38400, 57600, 115200}
        Dim baudRates() As Integer = {38400}
        RoofDeckBaudRateComboBox.Items.AddRange(baudRates.Cast(Of Object).ToArray())
        MidHeightBaudRateComboBox.Items.AddRange(baudRates.Cast(Of Object).ToArray())
        GroundFloorBaudRateComboBox.Items.AddRange(baudRates.Cast(Of Object).ToArray())

        ' Update the triggered data count
        UpdateTriggeredDataCount()

        UpdateStatusDisplay()

        ' Add items to CheckedListBox1
        CheckedListBox1.Items.Add("X", True)
        CheckedListBox1.Items.Add("Y", True)
        CheckedListBox1.Items.Add("Z", True)

        ' Add handler for ItemCheck event
        AddHandler CheckedListBox1.ItemCheck, AddressOf CheckedListBox1_ItemCheck

        'add export to string to textbox
        ExportToTextBox.Text = triggeredFolderPath

        SetupPlotView(GroundFloorPlotView, 1)
        SetupPlotView(MidHeightPlotView, 2)
        SetupPlotView(RoofDeckPlotView, 3)

        UpdateTimer.Interval = 10 ' Interval in milliseconds
        UpdateTimer.Start()

    End Sub

    Private Sub SetupPlotView(plotView As PlotView, floor As Integer)
        Dim timeSeries = New LineSeries
        Dim xSeries = New LineSeries
        Dim ySeries = New LineSeries
        Dim zSeries = New LineSeries

        Dim title = ""

        Select Case floor
            Case 1
                ' Initialize series for time, x, y, z
                timeSeries1 = New LineSeries() With {
                    .Title = "Time"
                }
                xSeries1 = New LineSeries() With {
                    .Title = "X"
                }
                ySeries1 = New LineSeries() With {
                    .Title = "Y"
                }
                zSeries1 = New LineSeries() With {
                    .Title = "Z"
                }

                timeSeries = timeSeries1
                xSeries = xSeries1
                ySeries = ySeries1
                zSeries = zSeries1

                title = "Ground Floor Data Plot"
            Case 2
                ' Initialize series for time, x, y, z
                timeSeries2 = New LineSeries() With {
                    .Title = "Time"
                }
                xSeries2 = New LineSeries() With {
                    .Title = "X"
                }
                ySeries2 = New LineSeries() With {
                    .Title = "Y"
                }
                zSeries2 = New LineSeries() With {
                    .Title = "Z"
                }

                timeSeries = timeSeries2
                xSeries = xSeries2
                ySeries = ySeries2
                zSeries = zSeries2

                title = "Mid Height Data Plot"
            Case 3
                ' Initialize series for time, x, y, z
                timeSeries3 = New LineSeries() With {
                    .Title = "Time"
                }
                xSeries3 = New LineSeries() With {
                    .Title = "X"
                }
                ySeries3 = New LineSeries() With {
                    .Title = "Y"
                }
                zSeries3 = New LineSeries() With {
                    .Title = "Z"
                }

                timeSeries = timeSeries3
                xSeries = xSeries3
                ySeries = ySeries3
                zSeries = zSeries3

                title = "Roof Deck Data Plot"
        End Select

        ' Create a new PlotModel
        Dim plotModel As New PlotModel() With {
            .Title = title
        }
        plotModel.Series.Add(timeSeries)
        plotModel.Series.Add(xSeries)
        plotModel.Series.Add(ySeries)
        plotModel.Series.Add(zSeries)


        ' Limit the amplitude of Y-axis for specific series
        Dim yAxisMin As Double = -20 ' Example minimum value
        Dim yAxisMax As Double = 20  ' Example maximum value

        ' Set the Y-axis limits for the series you want to limit
        ySeries.YAxisKey = "YAxisKey" ' Assign a key to the Y-axis
        plotModel.Axes.Add(New LinearAxis() With {
        .Key = "YAxisKey",
        .Position = AxisPosition.Left,
        .Minimum = yAxisMin,
        .Maximum = yAxisMax
    })

        ' Create a custom plot controller to disable zooming
        Dim customController As New PlotController()
        customController.UnbindMouseWheel() ' Disable zooming with the mouse wheel

        ' Assign the custom controller to the PlotView
        plotView.Controller = customController


        ' Set the Model property of PlotView to display the plotModel
        plotView.Model = plotModel
    End Sub

    Dim roofDeckArduinoData As String = ""
    Dim midHeightArduinoData As String = ""
    Dim groundFloorArduinoData As String = ""
    ' Methods to handle data received events for each serial port
    Private Sub SerialPort1_DataReceived(sender As Object, e As SerialDataReceivedEventArgs)
        Dim sp As SerialPort = CType(sender, SerialPort)
        Dim indata As String = sp.ReadLine()
        roofDeckArduinoData = indata
        ProcessData(indata, RoofDeckPlotView, 3, triggeredFilePath3, isTriggered3, triggerStartTime3, timeAfter20_3)
    End Sub

    Private Sub SerialPort2_DataReceived(sender As Object, e As SerialDataReceivedEventArgs)
        Dim sp As SerialPort = CType(sender, SerialPort)
        Dim indata As String = sp.ReadLine()
        midHeightArduinoData = indata
        ProcessData(indata, MidHeightPlotView, 2, triggeredFilePath2, isTriggered2, triggerStartTime2, timeAfter20_2)
    End Sub

    Private Sub SerialPort3_DataReceived(sender As Object, e As SerialDataReceivedEventArgs)
        Dim sp As SerialPort = CType(sender, SerialPort)
        Dim indata As String = sp.ReadLine()
        groundFloorArduinoData = indata
        ProcessData(indata, GroundFloorPlotView, 1, triggeredFilePath1, isTriggered1, triggerStartTime1, timeAfter20_1)
    End Sub

    ' Method to process incoming data
    ' Declare state variables for each Arduino
    Private triggeredFilePath1 As String
    Private triggeredFilePath2 As String
    Private triggeredFilePath3 As String

    Private isTriggered1 As Boolean = False
    Private isTriggered2 As Boolean = False
    Private isTriggered3 As Boolean = False

    Private triggerStartTime1 As DateTime
    Private triggerStartTime2 As DateTime
    Private triggerStartTime3 As DateTime

    'Private timeAfter20_1 As Double = 20.01
    'Private timeAfter20_2 As Double = 20.01
    'Private timeAfter20_3 As Double = 20.01
    Private timeAfter20_1 As Double = 5.01
    Private timeAfter20_2 As Double = 5.01
    Private timeAfter20_3 As Double = 5.01

    Private MaxPoints As Integer = 400  ' Maximum number of data points to display
    Private Sub ProcessData(data As String, plotView As PlotView, floor As Integer, ByRef triggeredFilePath As String, ByRef isTriggered As Boolean, ByRef triggerStartTime As DateTime, ByRef timeAfter20 As Double)
        Dim g_to_ms2 = 9.80665
        If isStarted Then
            Dim parts() As String = data.Split(" "c)
            If parts.Length = 4 Then
                Dim time As Double
                Dim x As Double
                Dim y As Double
                Dim z As Double

                If Double.TryParse(parts(0), time) AndAlso Double.TryParse(parts(1), x) AndAlso Double.TryParse(parts(2), y) AndAlso Double.TryParse(parts(3), z) Then


                    'Dim currentData As String = $"{timeAfter20.ToString("0.0000")} {x.ToString("0.0000")} {y.ToString("0.0000")} {z.ToString("0.0000")}"
                    Dim currentData As String = $"{timeAfter20.ToString("0.0000")}{vbTab}{x.ToString("0.0000")}{vbTab}{y.ToString("0.0000")}{vbTab}{z.ToString("0.0000")}"

                    Dim sensorName As String = ""

                    ' Check for trigger condition (x or y >= 0.05) for ground floor
                    If (Math.Abs(x) >= 0.05 * g_to_ms2 OrElse Math.Abs(y) >= 0.05 * g_to_ms2) AndAlso Not isTriggered Then
                        ' Start triggered recording
                        isTriggered = True
                        triggerStartTime = DateTime.Now
                        Select Case floor
                            Case 1
                                triggeredFilePath = Path.Combine(triggeredFolderPath, $"TriggeredData_GroundFloor_{DateTime.Now:yyyy-MM-dd-HH-mm-ss}.txt")
                                sensorName = "Ground Floor Sensor"
                            Case 2
                                triggeredFilePath = Path.Combine(triggeredFolderPath, $"TriggeredData_MidFloor_{DateTime.Now:yyyy-MM-dd-HH-mm-ss}.txt")
                                sensorName = "Mid Height Sensor"
                            Case 3
                                triggeredFilePath = Path.Combine(triggeredFolderPath, $"TriggeredData_RoofDeck_{DateTime.Now:yyyy-MM-dd-HH-mm-ss}.txt")
                                sensorName = "Roof Deck Sensor"
                        End Select

                        ' Start triggered recording and update status
                        UpdateStatus("Recording")

                        'add metadata
                        Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")
                        Dim currentTime As String = DateTime.Now.ToString("HH:mm")
                        Try
                            Using writer As StreamWriter = New StreamWriter(triggeredFilePath, True)
                                writer.WriteLine(sensorName)
                                writer.WriteLine("Date Recorded: " & currentDate)
                                writer.WriteLine("Time Recorded: " & currentTime)
                                writer.WriteLine("Frequency: 100hz")
                                writer.WriteLine("time[s]" & vbTab & "AccelX[m/s^2]" & vbTab & "AccelY[m/s^2]" & vbTab & "AccelZ[m/s^2]")

                            End Using

                        Catch ex As Exception
                            MessageBox.Show("Error recording file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                        ''add the first 20sec data
                        'Dim first20SecDataFilePath As String = Directory.GetCurrentDirectory() & "\First20SecData.txt"
                        'Try
                        '    ' Read all lines from the file into the list
                        '    Dim first20SecDataList As New List(Of String)(File.ReadAllLines(first20SecDataFilePath))

                        '    Using writer As StreamWriter = New StreamWriter(triggeredFilePath, True)
                        '        For Each datapoint As String In first20SecDataList
                        '            writer.WriteLine(datapoint)
                        '        Next
                        '    End Using

                        'Catch ex As Exception
                        '    MessageBox.Show("Error recording file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        'End Try
                        ' Add the first 20sec data
                        Dim first20SecDataFilePath As String = Directory.GetCurrentDirectory() & "\First20SecData.txt"
                        Dim maxRowsToRead As Integer = 501

                        Try
                            ' Read all lines from the file into the list
                            Dim first20SecDataList As New List(Of String)(File.ReadAllLines(first20SecDataFilePath))

                            Using writer As StreamWriter = New StreamWriter(triggeredFilePath, True)
                                For i As Integer = 0 To Math.Min(maxRowsToRead - 1, first20SecDataList.Count - 1)
                                    writer.WriteLine(first20SecDataList(i))
                                Next
                            End Using

                        Catch ex As Exception
                            MessageBox.Show("Error recording file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    End If

                    ' Check if triggered recording is active
                    If isTriggered AndAlso DateTime.Now.Subtract(triggerStartTime) <= triggerDuration Then

                        ' Write data to triggered file
                        Try
                            Using writer As StreamWriter = New StreamWriter(triggeredFilePath, True)
                                writer.WriteLine(currentData)
                                timeAfter20 += 0.01
                            End Using
                        Catch ex As Exception
                            'MessageBox.Show("Error writing triggered data to file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    ElseIf isTriggered Then

                        ' Stop triggered recording after duration
                        isTriggered = False
                        timeAfter20 = 20.01

                        'MessageBox.Show("Triggered recording stopped.", "Triggered Recording", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' Stop triggered recording and update status
                        UpdateStatus("Reading")

                        ' Update the triggered data count
                        UpdateTriggeredDataCount()

                        ' Correct net acceleration
                        CorrectNetAcceleration(triggeredFilePath)

                        triggeredFilePath = ""
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub CorrectSmallValues(filePath As String)
        Try
            ' Read all lines from the file
            Dim lines As List(Of String) = File.ReadAllLines(filePath).ToList()

            ' Separate metadata and data
            Dim metadataLines As List(Of String) = lines.Take(6).ToList()
            Dim dataLines As List(Of String) = lines.Skip(6).ToList()

            ' Process data to set small values to zero
            For i As Integer = 0 To dataLines.Count - 1
                Dim parts() As String = dataLines(i).Split(vbTab)
                If parts.Length = 4 Then
                    Dim time As Double = Double.Parse(parts(0))
                    Dim x As Double = Double.Parse(parts(1))
                    Dim y As Double = Double.Parse(parts(2))
                    Dim z As Double = Double.Parse(parts(3))

                    ' Set values close to zero to zero
                    If Math.Abs(x) <= 0.02 Then x = 0
                    If Math.Abs(y) <= 0.02 Then y = 0
                    If Math.Abs(z) <= 0.02 Then z = 0

                    dataLines(i) = $"{time.ToString("0.0000")}{vbTab}{x}{vbTab}{y}{vbTab}{z}"
                End If
            Next

            ' Combine metadata and processed data
            Dim finalLines As New List(Of String)(metadataLines)
            finalLines.AddRange(dataLines)

            ' Write the processed data back to the file
            File.WriteAllLines(filePath, finalLines)

        Catch ex As Exception
            MessageBox.Show("Error processing file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ApplyLowPassFilter(filePath As String)
        Try
            ' Read all lines from the file
            Dim lines As List(Of String) = File.ReadAllLines(filePath).ToList()

            ' Separate metadata and data
            Dim metadataLines As List(Of String) = lines.Take(6).ToList()
            Dim dataLines As List(Of String) = lines.Skip(6).ToList()

            ' Parameters for low-pass filter
            Dim alpha As Double = 0.1 ' Smoothing factor (0 < alpha < 1)
            Dim filteredX As Double = 0.0
            Dim filteredY As Double = 0.0
            Dim filteredZ As Double = 0.0

            ' Process data with low-pass filter
            For i As Integer = 0 To dataLines.Count - 1
                Dim parts() As String = dataLines(i).Split(vbTab)
                If parts.Length = 4 Then
                    Dim time As Double = Double.Parse(parts(0))
                    Dim x As Double = Double.Parse(parts(1))
                    Dim y As Double = Double.Parse(parts(2))
                    Dim z As Double = Double.Parse(parts(3))

                    ' Apply low-pass filter
                    filteredX = alpha * x + (1 - alpha) * filteredX
                    filteredY = alpha * y + (1 - alpha) * filteredY
                    filteredZ = alpha * z + (1 - alpha) * filteredZ

                    dataLines(i) = $"{time.ToString("0.0000")}{vbTab}{filteredX}{vbTab}{filteredY}{vbTab}{filteredZ}"
                End If
            Next

            ' Combine metadata and processed data
            Dim finalLines As New List(Of String)(metadataLines)
            finalLines.AddRange(dataLines)

            ' Write the processed data back to the file
            File.WriteAllLines(filePath, finalLines)

        Catch ex As Exception
            MessageBox.Show("Error processing file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Private Sub CorrectNetAcceleration(filePath As String)
    '    Try
    '        ''First, correct small values
    '        'CorrectSmallValues(filePath)

    '        ''Low-pass filter
    '        'ApplyLowPassFilter(filePath)

    '        ' Read all lines from the file
    '        Dim lines As List(Of String) = File.ReadAllLines(filePath).ToList()

    '        ' Separate metadata and data
    '        Dim metadataLines As List(Of String) = lines.Take(6).ToList()
    '        Dim dataLines As List(Of String) = lines.Skip(6).ToList()

    '        ' Find the index of the last instance where abs(x), abs(y), or abs(z) exceeds 0.05
    '        Dim lastIndex As Integer = dataLines.Count - 1
    '        For i As Integer = dataLines.Count - 1 To 0 Step -1
    '            Dim parts() As String = dataLines(i).Split(vbTab)
    '            If parts.Length = 4 Then
    '                Dim x, y, z As Double
    '                If Double.TryParse(parts(1), x) AndAlso Double.TryParse(parts(2), y) AndAlso Double.TryParse(parts(3), z) Then
    '                    If Math.Abs(x) >= 0.05 OrElse Math.Abs(y) >= 0.05 OrElse Math.Abs(z) >= 0.05 Then
    '                        lastIndex = i
    '                        Exit For
    '                    End If
    '                End If
    '            End If
    '        Next

    '        '' Skip the first 2001 samples after the metadata and consider samples only before lastIndex
    '        'Dim samplesToAdjust As List(Of String) = dataLines.Skip(2001).Take(lastIndex - 2001 + 1).ToList()
    '        ' Skip the first 501 samples after the metadata and consider samples only before lastIndex
    '        Dim samplesToAdjust As List(Of String) = dataLines.Skip(501).Take(lastIndex - 501 + 1).ToList()
    '        Dim remainingSamples As List(Of String) = dataLines.Skip(lastIndex + 1).ToList()

    '        ' Lists to hold x, y, z values for mean calculation
    '        Dim xValues As New List(Of Double)()
    '        Dim yValues As New List(Of Double)()
    '        Dim zValues As New List(Of Double)()

    '        ' Read values from the file
    '        For Each line As String In samplesToAdjust
    '            Dim parts() As String = line.Split(vbTab)
    '            If parts.Length = 4 Then
    '                Dim x, y, z As Double
    '                If Double.TryParse(parts(1), x) AndAlso Double.TryParse(parts(2), y) AndAlso Double.TryParse(parts(3), z) Then
    '                    xValues.Add(x)
    '                    yValues.Add(y)
    '                    zValues.Add(z)
    '                End If
    '            End If
    '        Next

    '        ' Calculate net acceleration
    '        Dim netX As Double = xValues.Sum()
    '        Dim netY As Double = yValues.Sum()
    '        Dim netZ As Double = zValues.Sum()

    '        ' Calculate adjustments
    '        Dim adjustmentX As Double = netX / samplesToAdjust.Count
    '        Dim adjustmentY As Double = netY / samplesToAdjust.Count
    '        Dim adjustmentZ As Double = netZ / samplesToAdjust.Count

    '        ' Apply adjustments to zero net acceleration and set small values to zero
    '        For i As Integer = 0 To samplesToAdjust.Count - 1
    '            Dim parts() As String = samplesToAdjust(i).Split(vbTab)
    '            If parts.Length = 4 Then
    '                Dim time As Double = Double.Parse(parts(0))
    '                Dim x As Double = Double.Parse(parts(1)) - adjustmentX
    '                Dim y As Double = Double.Parse(parts(2)) - adjustmentY
    '                Dim z As Double = Double.Parse(parts(3)) - adjustmentZ

    '                samplesToAdjust(i) = $"{time.ToString("0.0000")}{vbTab}{x}{vbTab}{y}{vbTab}{z}"
    '            End If
    '        Next

    '        ' Set small values in remainingSamples to zero
    '        For i As Integer = 0 To remainingSamples.Count - 1
    '            Dim parts() As String = remainingSamples(i).Split(vbTab)
    '            If parts.Length = 4 Then
    '                Dim time As Double = Double.Parse(parts(0))
    '                Dim x As Double = Double.Parse(parts(1))
    '                Dim y As Double = Double.Parse(parts(2))
    '                Dim z As Double = Double.Parse(parts(3))

    '                If Math.Abs(x) <= 0.03 Then x = 0
    '                If Math.Abs(y) <= 0.03 Then y = 0
    '                If Math.Abs(z) <= 0.03 Then z = 0

    '                remainingSamples(i) = $"{time.ToString("0.0000")}{vbTab}{x}{vbTab}{y}{vbTab}{z}"
    '            End If
    '        Next

    '        ' Combine metadata, first 2001 samples, adjusted samples, and remaining samples
    '        Dim finalLines As New List(Of String)(metadataLines)
    '        'finalLines.AddRange(dataLines.Take(2001))
    '        finalLines.AddRange(dataLines.Take(501))
    '        finalLines.AddRange(samplesToAdjust)
    '        finalLines.AddRange(remainingSamples)

    '        ' Write the adjusted data back to the file
    '        File.WriteAllLines(filePath, finalLines)

    '    Catch ex As Exception
    '        MessageBox.Show("Error processing file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Private Sub CorrectNetAcceleration(filePath As String)
        Dim g_to_ms2 = 9.80665
        Try
            ' Read all lines from the file
            Dim lines As List(Of String) = File.ReadAllLines(filePath).ToList()

            ' Separate metadata and data
            Dim metadataLines As List(Of String) = lines.Take(6).ToList()
            Dim dataLines As List(Of String) = lines.Skip(6).ToList()

            ' Find the index of the last instance where abs(x), abs(y), or abs(z) exceeds 0.05
            Dim lastIndex As Integer = dataLines.Count - 1
            For i As Integer = dataLines.Count - 1 To 0 Step -1
                Dim parts() As String = dataLines(i).Split(vbTab)
                If parts.Length = 4 Then
                    Dim x, y, z As Double
                    If Double.TryParse(parts(1), x) AndAlso Double.TryParse(parts(2), y) AndAlso Double.TryParse(parts(3), z) Then
                        If Math.Abs(x) >= 0.05 * g_to_ms2 OrElse Math.Abs(y) >= 0.05 * g_to_ms2 OrElse Math.Abs(z) >= 0.05 * g_to_ms2 Then
                            lastIndex = i
                            Exit For
                        End If
                    End If
                End If
            Next

            ' Skip the first 501 samples after the metadata and consider samples only before lastIndex
            Dim samplesToAdjust As List(Of String) = dataLines.Skip(501).Take(lastIndex - 501 + 1).ToList()
            Dim remainingSamples As List(Of String) = dataLines.Skip(lastIndex + 1).ToList()

            ' Lists to hold x, y, z values for mean calculation
            Dim xValues As New List(Of Double)()
            Dim yValues As New List(Of Double)()
            Dim zValues As New List(Of Double)()

            ' Read values from the file
            For Each line As String In samplesToAdjust
                Dim parts() As String = line.Split(vbTab)
                If parts.Length = 4 Then
                    Dim x, y, z As Double
                    If Double.TryParse(parts(1), x) AndAlso Double.TryParse(parts(2), y) AndAlso Double.TryParse(parts(3), z) Then
                        xValues.Add(x)
                        yValues.Add(y)
                        zValues.Add(z)
                    End If
                End If
            Next

            ' Calculate net acceleration
            Dim netX As Double = xValues.Sum()
            Dim netY As Double = yValues.Sum()
            Dim netZ As Double = zValues.Sum()

            ' Calculate adjustments
            Dim adjustmentX As Double = netX / samplesToAdjust.Count
            Dim adjustmentY As Double = netY / samplesToAdjust.Count
            Dim adjustmentZ As Double = netZ / samplesToAdjust.Count

            ' Apply adjustments to zero net acceleration
            For i As Integer = 0 To samplesToAdjust.Count - 1
                Dim parts() As String = samplesToAdjust(i).Split(vbTab)
                If parts.Length = 4 Then
                    Dim time As Double = Double.Parse(parts(0))
                    Dim x As Double = Double.Parse(parts(1)) - adjustmentX
                    Dim y As Double = Double.Parse(parts(2)) - adjustmentY
                    Dim z As Double = Double.Parse(parts(3)) - adjustmentZ

                    samplesToAdjust(i) = $"{time.ToString("0.0000")}{vbTab}{x}{vbTab}{y}{vbTab}{z}"
                End If
            Next

            ' Reread values after initial adjustments
            xValues.Clear()
            yValues.Clear()
            zValues.Clear()

            For Each line As String In samplesToAdjust
                Dim parts() As String = line.Split(vbTab)
                If parts.Length = 4 Then
                    Dim x, y, z As Double
                    If Double.TryParse(parts(1), x) AndAlso Double.TryParse(parts(2), y) AndAlso Double.TryParse(parts(3), z) Then
                        xValues.Add(x)
                        yValues.Add(y)
                        zValues.Add(z)
                    End If
                End If
            Next

            ' Calculate remaining net acceleration
            netX = xValues.Sum()
            netY = yValues.Sum()
            netZ = zValues.Sum()

            ' Adjust the last value to make net acceleration exactly zero
            If samplesToAdjust.Count > 0 Then
                Dim lastIndexAdjust As Integer = samplesToAdjust.Count - 1
                Dim lastParts() As String = samplesToAdjust(lastIndexAdjust).Split(vbTab)
                If lastParts.Length = 4 Then
                    Dim time As Double = Double.Parse(lastParts(0))
                    Dim x As Double = Double.Parse(lastParts(1)) - (netX)
                    Dim y As Double = Double.Parse(lastParts(2)) - (netY)
                    Dim z As Double = Double.Parse(lastParts(3)) - (netZ)

                    samplesToAdjust(lastIndexAdjust) = $"{time.ToString("0.0000")}{vbTab}{x}{vbTab}{y}{vbTab}{z}"
                End If
            End If

            ' Set small values in remainingSamples to zero
            For i As Integer = 0 To remainingSamples.Count - 1
                Dim parts() As String = remainingSamples(i).Split(vbTab)
                If parts.Length = 4 Then
                    Dim time As Double = Double.Parse(parts(0))
                    Dim x As Double = Double.Parse(parts(1))
                    Dim y As Double = Double.Parse(parts(2))
                    Dim z As Double = Double.Parse(parts(3))

                    If Math.Abs(x) <= 0.05 * g_to_ms2 Then x = 0
                    If Math.Abs(y) <= 0.05 * g_to_ms2 Then y = 0
                    If Math.Abs(z) <= 0.05 * g_to_ms2 Then z = 0

                    remainingSamples(i) = $"{time.ToString("0.0000")}{vbTab}{x}{vbTab}{y}{vbTab}{z}"
                End If
            Next

            ' Combine metadata, first 501 samples, adjusted samples, and remaining samples
            Dim finalLines As New List(Of String)(metadataLines)
            finalLines.AddRange(dataLines.Take(501))
            finalLines.AddRange(samplesToAdjust)
            'finalLines.AddRange(remainingSamples)

            ' Write the adjusted data back to the file
            File.WriteAllLines(filePath, finalLines)

        Catch ex As Exception
            MessageBox.Show("Error processing file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub







    Private Sub UpdatePlotOnUIThread(time As Double, x As Double, y As Double, z As Double,
                                 xSeries As LineSeries, ySeries As LineSeries, zSeries As LineSeries,
                                 timeSeries As LineSeries, MaxPoints As Integer,
                                 PlotView1 As OxyPlot.WindowsForms.PlotView)
        ' Update plot on the UI thread
        Me.Invoke(Sub()
                      ' Add new data points
                      xSeries.Points.Add(New DataPoint(time, x))
                      ySeries.Points.Add(New DataPoint(time, y))
                      zSeries.Points.Add(New DataPoint(time, z))

                      ' Remove old points if needed
                      If xSeries.Points.Count > MaxPoints Then
                          'timeSeries.Points.RemoveAt(0)
                          xSeries.Points.RemoveAt(0)
                          ySeries.Points.RemoveAt(0)
                          zSeries.Points.RemoveAt(0)
                      End If

                      ' Refresh the plot view
                      PlotView1.InvalidatePlot(True)
                  End Sub)
    End Sub



    ' This method handles the Connect button click event
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ConnectButton.Click
        ' Connect to selected COM ports
        Try
            Dim isConnected As Boolean = False

            ' Close existing ports
            For Each sp As SerialPort In serialPorts
                If sp IsNot Nothing AndAlso sp.IsOpen Then
                    sp.Close()
                End If
            Next

            'clear all listboxes
            RoofDeckListBox.Items.Clear()
            MidHeightListBox.Items.Clear()
            GroundFloorListBox.Items.Clear()

            ' Open new ports
            If RoofDeckComPortComboBox.SelectedIndex <> -1 AndAlso (RoofDeckBaudRateComboBox.SelectedIndex <> -1 Or RoofDeckBaudRateComboBox.Text <> "") Then
                serialPorts(0) = New SerialPort(RoofDeckComPortComboBox.Text, CInt(RoofDeckBaudRateComboBox.Text))
                serialPorts(0).Open()
                Thread.Sleep(1000) ' Delay after opening port
                AddHandler serialPorts(0).DataReceived, AddressOf SerialPort1_DataReceived
                isConnected = True

                RoofDeckListBox.Items.Add("Roof deck sensor calibrated.")
            End If
            If MidHeightComPortComboBox.SelectedIndex <> -1 AndAlso (MidHeightBaudRateComboBox.SelectedIndex <> -1 Or MidHeightBaudRateComboBox.Text <> "") Then
                serialPorts(1) = New SerialPort(MidHeightComPortComboBox.Text, CInt(MidHeightBaudRateComboBox.Text))
                serialPorts(1).Open()
                Thread.Sleep(1000) ' Delay after opening port
                AddHandler serialPorts(1).DataReceived, AddressOf SerialPort2_DataReceived
                isConnected = True

                MidHeightListBox.Items.Add("Mid height sensor calibrated.")
            End If
            If GroundFloorComPortComboBox.SelectedIndex <> -1 AndAlso (GroundFloorBaudRateComboBox.SelectedIndex <> -1 Or GroundFloorBaudRateComboBox.Text <> "") Then
                serialPorts(2) = New SerialPort(GroundFloorComPortComboBox.Text, CInt(GroundFloorBaudRateComboBox.Text))
                serialPorts(2).Open()
                Thread.Sleep(1000) ' Delay after opening port
                AddHandler serialPorts(2).DataReceived, AddressOf SerialPort3_DataReceived
                isConnected = True

                GroundFloorListBox.Items.Add("Ground floor sensor calibrated.")
            End If

            If isConnected Then
                MessageBox.Show("COM ports connected successfully!", "Connection Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("No COM ports were selected.", "Connection Status", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' This method handles the Start button click event
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles StartButton.Click
        isStarted = True ' Set the flag to true
        UpdateStatusDisplay()
        ClearPlot()
        Try
            ' Clear existing files if they exist when starting recording
            If File.Exists(filePath) Then
                File.Delete(filePath)
            End If
            If File.Exists(triggeredFilePath) Then
                File.Delete(triggeredFilePath)
            End If

            ' Open serial ports and send "start" command
            For Each sp In serialPorts
                If sp IsNot Nothing AndAlso sp.IsOpen Then
                    sp.WriteLine("start")
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error starting recording: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' This method handles the End button click event
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles EndButton.Click
        isStarted = False ' Set the flag to false
        isTriggered = False
        UpdateStatusDisplay()
        For Each sp In serialPorts
            If sp IsNot Nothing AndAlso sp.IsOpen Then
                sp.WriteLine("end")
            End If
        Next
    End Sub

    ' This method handles the Clear button click event
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        ClearPlot()
    End Sub

    Private Sub ClearPlot()
        ' Clear data points from all series
        timeSeries1.Points.Clear()
        xSeries1.Points.Clear()
        ySeries1.Points.Clear()
        zSeries1.Points.Clear()
        timeSeries2.Points.Clear()
        xSeries2.Points.Clear()
        ySeries2.Points.Clear()
        zSeries2.Points.Clear()
        timeSeries3.Points.Clear()
        xSeries3.Points.Clear()
        ySeries3.Points.Clear()
        zSeries3.Points.Clear()
        ' Refresh the plot view
        GroundFloorPlotView.InvalidatePlot(True)
        MidHeightPlotView.InvalidatePlot(True)
        RoofDeckPlotView.InvalidatePlot(True)
    End Sub

    Private Sub UpdateTriggeredDataCount()
        Try
            ' Get the desktop path
            Dim triggedDataFilePath As String = triggeredFolderPath

            ' Search for files with the pattern "TriggeredData_*.txt" on the desktop
            Dim triggeredFiles As String() = Directory.GetFiles(triggedDataFilePath, "TriggeredData_*.txt")

            ' Update the count in Label5
            Invoke(Sub()
                       FileCountLabel.Text = triggeredFiles.Length.ToString()
                   End Sub)
        Catch ex As Exception
            MessageBox.Show("Error counting triggered data files: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UpdateStatus(status As String)
        ' Use Invoke if required to update UI controls from a background thread
        If StatusTextBox.InvokeRequired Then
            StatusTextBox.Invoke(Sub() UpdateStatus(status))
        Else
            StatusTextBox.Text = status
            ' Update EndButton status
            EndButton.Enabled = True
            If status = "Recording" Then
                EndButton.Enabled = False
            End If
        End If
    End Sub

    Private Sub UpdateStatusDisplay()
        If Not isStarted Then
            UpdateStatus("Standby")
        ElseIf isTriggered Then
            UpdateStatus("Recording")
        Else
            UpdateStatus("Reading")
        End If
    End Sub

    Private Sub OpenFolderButton_Click(sender As Object, e As EventArgs)
        Try
            If Directory.Exists(triggeredFolderPath) Then
                Process.Start("explorer.exe", triggeredFolderPath)
            Else
                MessageBox.Show("Triggered data folder does not exist.", "Folder Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show("Error opening folder: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        ' Get the item that is being checked/unchecked
        Dim item As String = CheckedListBox1.Items(e.Index).ToString()

        ' Determine the new checked state
        Dim isChecked As Boolean = (e.NewValue = CheckState.Checked)

        ' Toggle the visibility of the corresponding series in all charts
        Select Case item
            Case "X"
                xSeries1.IsVisible = isChecked
                xSeries2.IsVisible = isChecked
                xSeries3.IsVisible = isChecked
                GroundFloorPlotView.InvalidatePlot(True) ' Refresh the plot view
                MidHeightPlotView.InvalidatePlot(True) ' Refresh the plot view
                RoofDeckPlotView.InvalidatePlot(True) ' Refresh the plot view
            Case "Y"
                ySeries1.IsVisible = isChecked
                ySeries2.IsVisible = isChecked
                ySeries3.IsVisible = isChecked
                GroundFloorPlotView.InvalidatePlot(True) ' Refresh the plot view
                MidHeightPlotView.InvalidatePlot(True) ' Refresh the plot view
                RoofDeckPlotView.InvalidatePlot(True) ' Refresh the plot view
            Case "Z"
                zSeries1.IsVisible = isChecked
                zSeries2.IsVisible = isChecked
                zSeries3.IsVisible = isChecked
                GroundFloorPlotView.InvalidatePlot(True) ' Refresh the plot view
                MidHeightPlotView.InvalidatePlot(True) ' Refresh the plot view
                RoofDeckPlotView.InvalidatePlot(True) ' Refresh the plot view
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles RefreshButton.Click
        RoofDeckComPortComboBox.Items.Clear()
        MidHeightComPortComboBox.Items.Clear()
        GroundFloorComPortComboBox.Items.Clear()
        For Each port_name As String In IO.Ports.SerialPort.GetPortNames
            RoofDeckComPortComboBox.Items.Add(port_name)
            MidHeightComPortComboBox.Items.Add(port_name)
            GroundFloorComPortComboBox.Items.Add(port_name)
        Next
    End Sub

    Private Sub ExportToButton_Click(sender As Object, e As EventArgs) Handles ExportToButton.Click
        ' Create a FolderBrowserDialog instance
        Dim folderDialog As New FolderBrowserDialog()

        ' Set initial directory if needed
        ' folderDialog.SelectedPath = "C:\Initial\Directory"

        ' Show the FolderBrowserDialog and check if the user selected a folder
        If folderDialog.ShowDialog() = DialogResult.OK Then
            ' Assign the selected folder path to triggeredFolderPath
            triggeredFolderPath = folderDialog.SelectedPath

            ' Optionally, display the selected path or perform other actions
            MessageBox.Show($"Selected Folder: {triggeredFolderPath}", "Folder Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ExportToTextBox.Text = triggeredFolderPath
        Else
            ' Optionally, handle case where the user cancels the dialog
            MessageBox.Show("Operation cancelled by user.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub


    Private Sub updateTimer_Tick(sender As Object, e As EventArgs) Handles UpdateTimer.Tick
        If isStarted Then
            Dim data3 = roofDeckArduinoData

            Dim parts3() As String = data3.Split(" "c)
            If parts3.Length = 4 Then
                Dim time As Double
                Dim x As Double
                Dim y As Double
                Dim z As Double

                If Double.TryParse(parts3(0), time) AndAlso Double.TryParse(parts3(1), x) AndAlso Double.TryParse(parts3(2), y) AndAlso Double.TryParse(parts3(3), z) Then
                    UpdatePlotOnUIThread(time, x, y, z, xSeries3, ySeries3, zSeries3, timeSeries3, MaxPoints, RoofDeckPlotView)
                End If
            End If

            Dim data2 = midHeightArduinoData

            Dim parts2() As String = data2.Split(" "c)
            If parts2.Length = 4 Then
                Dim time As Double
                Dim x As Double
                Dim y As Double
                Dim z As Double

                If Double.TryParse(parts2(0), time) AndAlso Double.TryParse(parts2(1), x) AndAlso Double.TryParse(parts2(2), y) AndAlso Double.TryParse(parts2(3), z) Then
                    UpdatePlotOnUIThread(time, x, y, z, xSeries2, ySeries2, zSeries2, timeSeries2, MaxPoints, MidHeightPlotView)
                End If
            End If
        End If

        Dim data1 = groundFloorArduinoData

        Dim parts1() As String = data1.Split(" "c)
        If parts1.Length = 4 Then
            Dim time As Double
            Dim x As Double
            Dim y As Double
            Dim z As Double

            If Double.TryParse(parts1(0), time) AndAlso Double.TryParse(parts1(1), x) AndAlso Double.TryParse(parts1(2), y) AndAlso Double.TryParse(parts1(3), z) Then
                UpdatePlotOnUIThread(time, x, y, z, xSeries1, ySeries1, zSeries1, timeSeries1, MaxPoints, GroundFloorPlotView)
            End If
        End If
    End Sub

    Private Sub FileRecordedButton_Click(sender As Object, e As EventArgs) Handles FileRecordedButton.Click
        Try
            If Directory.Exists(triggeredFolderPath) Then
                Process.Start("explorer.exe", triggeredFolderPath)
            Else
                MessageBox.Show("Triggered data folder does not exist.", "Folder Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show("Error opening folder: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


#Region "Data Conversion"
    Private Sub SelectRecordingButton_Click(sender As Object, e As EventArgs) Handles SelectRecordingButton.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            LoadDataAndPlot(filePath)

            ' Get only the filename and add it to the RecordingTitleListBox
            RecordingTitleListBox.Items.Clear()
            Dim fileName As String = Path.GetFileName(filePath)
            RecordingTitleListBox.Items.Add(fileName)
        End If
    End Sub

    Private Sub LoadDataAndPlot(filePath As String)
        ' Create PlotModels for X, Y, and Z accelerations, velocities, and displacements
        Dim plotModelXAccel As New PlotModel With {.Title = "X Acceleration vs Time"}
        Dim plotModelYAccel As New PlotModel With {.Title = "Y Acceleration vs Time"}
        Dim plotModelZAccel As New PlotModel With {.Title = "Z Acceleration vs Time"}

        Dim plotModelXVel As New PlotModel With {.Title = "X Velocity vs Time"}
        Dim plotModelYVel As New PlotModel With {.Title = "Y Velocity vs Time"}
        Dim plotModelZVel As New PlotModel With {.Title = "Z Velocity vs Time"}

        Dim plotModelXDisp As New PlotModel With {.Title = "X Displacement vs Time"}
        Dim plotModelYDisp As New PlotModel With {.Title = "Y Displacement vs Time"}
        Dim plotModelZDisp As New PlotModel With {.Title = "Z Displacement vs Time"}

        ' Create LineSeries for the X, Y, and Z Accelerations, Velocities, and Displacements without markers and with blue color
        Dim seriesXAccel As New LineSeries With {.Title = "AccelX", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}
        Dim seriesYAccel As New LineSeries With {.Title = "AccelY", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}
        Dim seriesZAccel As New LineSeries With {.Title = "AccelZ", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}

        Dim seriesXVel As New LineSeries With {.Title = "VelX", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}
        Dim seriesYVel As New LineSeries With {.Title = "VelY", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}
        Dim seriesZVel As New LineSeries With {.Title = "VelZ", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}

        Dim seriesXDisp As New LineSeries With {.Title = "DispX", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}
        Dim seriesYDisp As New LineSeries With {.Title = "DispY", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}
        Dim seriesZDisp As New LineSeries With {.Title = "DispZ", .MarkerType = MarkerType.None, .Color = OxyColors.Blue}

        ' Clear existing rows in the DataGridViews
        XAccelerationDGV.Rows.Clear()
        YAccelerationDGV.Rows.Clear()
        ZAccelerationDGV.Rows.Clear()
        XVelocityDGV.Rows.Clear()
        YVelocityDGV.Rows.Clear()
        ZVelocityDGV.Rows.Clear()
        XDisplacementDGV.Rows.Clear()
        YDisplacementDGV.Rows.Clear()
        ZDisplacementDGV.Rows.Clear()

        ' Initialize variables to track max values and their corresponding times
        Dim maxAccelX, maxAccelY, maxAccelZ As Double
        Dim maxAccelXTime, maxAccelYTime, maxAccelZTime As Double

        Dim maxVelX, maxVelY, maxVelZ As Double
        Dim maxVelXTime, maxVelYTime, maxVelZTime As Double

        Dim maxDispX, maxDispY, maxDispZ As Double
        Dim maxDispXTime, maxDispYTime, maxDispZTime As Double

        ' Read the file and parse data
        Dim lines() As String = File.ReadAllLines(filePath)
        Dim deltaTime As Double = 0.01 ' Assuming a fixed time interval (e.g., 0.01 seconds, you can modify as needed)

        ' Variables to hold the running sums for velocity and displacement
        Dim velocityX, velocityY, velocityZ As Double
        Dim displacementX, displacementY, displacementZ As Double

        Dim time As Double = 0
        Try
            ' Skip the first 6 lines of metadata
            For i As Integer = 6 To lines.Length - 2 ' n-1 because we need i+1 data
                Dim parts() As String = lines(i).Split(vbTab)
                'Dim time As Double = Double.Parse(parts(0))
                time += 0.01
                Dim accelX As Double = Double.Parse(parts(1))
                Dim accelY As Double = Double.Parse(parts(2))
                Dim accelZ As Double = Double.Parse(parts(3))

                Dim nextParts() As String = lines(i + 1).Split(vbTab)
                Dim nextAccelX As Double = Double.Parse(nextParts(1))
                Dim nextAccelY As Double = Double.Parse(nextParts(2))
                Dim nextAccelZ As Double = Double.Parse(nextParts(3))

                ' Update max acceleration values and times
                If Math.Abs(accelX) > maxAccelX Then
                    maxAccelX = Math.Abs(accelX)
                    maxAccelXTime = time
                End If
                If Math.Abs(accelY) > maxAccelY Then
                    maxAccelY = Math.Abs(accelY)
                    maxAccelYTime = time
                End If
                If Math.Abs(accelZ) > maxAccelZ Then
                    maxAccelZ = Math.Abs(accelZ)
                    maxAccelZTime = time
                End If

                ' Calculate the change in velocity using the provided formula
                Dim dVX As Double = 0.5 * (accelX + nextAccelX) * deltaTime
                Dim dVY As Double = 0.5 * (accelY + nextAccelY) * deltaTime
                Dim dVZ As Double = 0.5 * (accelZ + nextAccelZ) * deltaTime

                velocityX += dVX
                velocityY += dVY
                velocityZ += dVZ

                ' Update max velocity values and times
                If Math.Abs(velocityX) > maxVelX Then
                    maxVelX = Math.Abs(velocityX)
                    maxVelXTime = time
                End If
                If Math.Abs(velocityY) > maxVelY Then
                    maxVelY = Math.Abs(velocityY)
                    maxVelYTime = time
                End If
                If Math.Abs(velocityZ) > maxVelZ Then
                    maxVelZ = Math.Abs(velocityZ)
                    maxVelZTime = time
                End If

                ' Trapezoidal integration to calculate displacement
                Dim dDX As Double = (velocityX + (velocityX - dVX)) * (deltaTime / 2)
                Dim dDY As Double = (velocityY + (velocityY - dVY)) * (deltaTime / 2)
                Dim dDZ As Double = (velocityZ + (velocityZ - dVZ)) * (deltaTime / 2)

                displacementX += dDX
                displacementY += dDY
                displacementZ += dDZ

                ' Update max displacement values and times
                If Math.Abs(displacementX) > maxDispX Then
                    maxDispX = Math.Abs(displacementX)
                    maxDispXTime = time
                End If
                If Math.Abs(displacementY) > maxDispY Then
                    maxDispY = Math.Abs(displacementY)
                    maxDispYTime = time
                End If
                If Math.Abs(displacementZ) > maxDispZ Then
                    maxDispZ = Math.Abs(displacementZ)
                    maxDispZTime = time
                End If

                ' Add points to the respective series
                seriesXAccel.Points.Add(New DataPoint(time.ToString("F2"), accelX.ToString("F5")))
                seriesYAccel.Points.Add(New DataPoint(time.ToString("F2"), accelY.ToString("F5")))
                seriesZAccel.Points.Add(New DataPoint(time.ToString("F2"), accelZ.ToString("F5")))

                seriesXVel.Points.Add(New DataPoint(time.ToString("F2"), velocityX.ToString("F5")))
                seriesYVel.Points.Add(New DataPoint(time.ToString("F2"), velocityY.ToString("F5")))
                seriesZVel.Points.Add(New DataPoint(time.ToString("F2"), velocityZ.ToString("F5")))

                seriesXDisp.Points.Add(New DataPoint(time.ToString("F2"), displacementX.ToString("F5")))
                seriesYDisp.Points.Add(New DataPoint(time.ToString("F2"), displacementY.ToString("F5")))
                seriesZDisp.Points.Add(New DataPoint(time.ToString("F2"), displacementZ.ToString("F5")))

                ' Add the data to the respective DataGridViews
                XAccelerationDGV.Rows.Add(time.ToString("F5"), accelX.ToString("F5"))
                YAccelerationDGV.Rows.Add(time.ToString("F5"), accelY.ToString("F5"))
                ZAccelerationDGV.Rows.Add(time.ToString("F5"), accelZ.ToString("F5"))
                XVelocityDGV.Rows.Add(time.ToString("F5"), velocityX.ToString("F5"))
                YVelocityDGV.Rows.Add(time.ToString("F5"), velocityY.ToString("F5"))
                ZVelocityDGV.Rows.Add(time.ToString("F5"), velocityZ.ToString("F5"))
                XDisplacementDGV.Rows.Add(time.ToString("F5"), displacementX.ToString("F5"))
                YDisplacementDGV.Rows.Add(time.ToString("F5"), displacementY.ToString("F5"))
                ZDisplacementDGV.Rows.Add(time.ToString("F5"), displacementZ.ToString("F5"))
            Next
        Catch ex As Exception
            MessageBox.Show("Error reading file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Display max values and times in the text boxes
        XAccelMaxTextBox.Text = $"{maxAccelX.ToString("F5")}m/s^2 @ {maxAccelXTime.ToString("F2")}s"
        YAccelMaxTextBox.Text = $"{maxAccelY.ToString("F5")}m/s^2 @ {maxAccelYTime.ToString("F2")}s"
        ZAccelMaxTextBox.Text = $"{maxAccelZ.ToString("F5")}m/s^2 @ {maxAccelZTime.ToString("F2")}s"

        XVelMaxTextBox.Text = $"{maxVelX.ToString("F5")}m/s @ {maxVelXTime.ToString("F2")}s"
        YVelMaxTextBox.Text = $"{maxVelY.ToString("F5")}m/s @ {maxVelYTime.ToString("F2")}s"
        ZVelMaxTextBox.Text = $"{maxVelZ.ToString("F5")}m/s @ {maxVelZTime.ToString("F2")}s"

        XDispMaxTextBox.Text = $"{maxDispX.ToString("F5")}m @ {maxDispXTime.ToString("F2")}s"
        YDispMaxTextBox.Text = $"{maxDispY.ToString("F5")}m @ {maxDispYTime.ToString("F2")}s"
        ZDispMaxTextBox.Text = $"{maxDispZ.ToString("F5")}m @ {maxDispZTime.ToString("F2")}s"

        ' Add the series to their respective PlotModels
        plotModelXAccel.Series.Add(seriesXAccel)
        plotModelYAccel.Series.Add(seriesYAccel)
        plotModelZAccel.Series.Add(seriesZAccel)

        plotModelXVel.Series.Add(seriesXVel)
        plotModelYVel.Series.Add(seriesYVel)
        plotModelZVel.Series.Add(seriesZVel)

        plotModelXDisp.Series.Add(seriesXDisp)
        plotModelYDisp.Series.Add(seriesYDisp)
        plotModelZDisp.Series.Add(seriesZDisp)

        ' Assign the PlotModels to the respective PlotViews
        XAccelerationPlotView.Model = plotModelXAccel
        YAccelerationPlotView.Model = plotModelYAccel
        ZAccelerationPlotView.Model = plotModelZAccel

        XVelocityPlotView.Model = plotModelXVel
        YVelocityPlotView.Model = plotModelYVel
        ZVelocityPlotView.Model = plotModelZVel

        XDisplacementPlotView.Model = plotModelXDisp
        YDisplacementPlotView.Model = plotModelYDisp
        ZDisplacementPlotView.Model = plotModelZDisp
    End Sub

    Private Sub SaveResultsButton_Click(sender As Object, e As EventArgs) Handles SaveResultsButton.Click
        ' Open a SaveFileDialog to choose where to save the file
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        saveFileDialog.Title = "Save Results As"
        saveFileDialog.FileName = "Results.txt"

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = saveFileDialog.FileName

            Try
                ' Open the file to write
                Using writer As New StreamWriter(filePath)
                    ' Write headers
                    'writer.WriteLine("Time[s]  AccelX[m/s²]  VelX[m/s]  DispX[m]  AccelY[m/s²]  VelY[m/s]  DispY[m]  AccelZ[m/s²]  VelZ[m/s]  DispZ[m]")
                    writer.WriteLine("Time[s]" & vbTab & "AccelX[m/s²]" & vbTab & "VelX[m/s]" & vbTab & "DispX[m]" & vbTab & "AccelY[m/s²]" & vbTab & "VelY[m/s]" & vbTab & "DispY[m]" & vbTab & "AccelZ[m/s²]" & vbTab & "VelZ[m/s]" & vbTab & "DispZ[m]")


                    ' Write data from DataGridViews
                    For i As Integer = 0 To XAccelerationDGV.Rows.Count - 1
                        Dim time As String = XAccelerationDGV.Rows(i).Cells(0).Value.ToString()

                        Dim accelX As String = XAccelerationDGV.Rows(i).Cells(1).Value.ToString()
                        Dim velX As String = XVelocityDGV.Rows(i).Cells(1).Value.ToString()
                        Dim dispX As String = XDisplacementDGV.Rows(i).Cells(1).Value.ToString()

                        Dim accelY As String = YAccelerationDGV.Rows(i).Cells(1).Value.ToString()
                        Dim velY As String = YVelocityDGV.Rows(i).Cells(1).Value.ToString()
                        Dim dispY As String = YDisplacementDGV.Rows(i).Cells(1).Value.ToString()

                        Dim accelZ As String = ZAccelerationDGV.Rows(i).Cells(1).Value.ToString()
                        Dim velZ As String = ZVelocityDGV.Rows(i).Cells(1).Value.ToString()
                        Dim dispZ As String = ZDisplacementDGV.Rows(i).Cells(1).Value.ToString()

                        ' Write the formatted line to the file
                        'writer.WriteLine($"{time}  {accelX}  {velX}  {dispX}  {accelY}  {velY}  {dispY}  {accelZ}  {velZ}  {dispZ}")
                        writer.WriteLine($"{time}{vbTab}{accelX}{vbTab}{velX}{vbTab}{dispX}{vbTab}{accelY}{vbTab}{velY}{vbTab}{dispY}{vbTab}{accelZ}{vbTab}{velZ}{vbTab}{dispZ}")

                    Next
                End Using

                MessageBox.Show("Results saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error saving results: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

#End Region

#Region "Visualization"
    Private groundXDisplacementData As List(Of Double)
    Private midHeightXDisplacementData As List(Of Double)
    Private roofXDisplacementData As List(Of Double)
    Private groundYDisplacementData As List(Of Double)
    Private midHeightYDisplacementData As List(Of Double)
    Private roofYDisplacementData As List(Of Double)
    Private groundZDisplacementData As List(Of Double)
    Private midHeightZDisplacementData As List(Of Double)
    Private roofZDisplacementData As List(Of Double)

    Private groundCurrentIndex As Integer = 0
    Private midHeightCurrentIndex As Integer = 0
    Private roofCurrentIndex As Integer = 0

    Private groundCirclePositionX As Integer = 0
    Private midHeightCirclePositionX As Integer = 0
    Private roofCirclePositionX As Integer = 0
    Private groundCirclePositionY As Integer = 0
    Private midHeightCirclePositionY As Integer = 0
    Private roofCirclePositionY As Integer = 0
    Private groundCirclePositionZ As Integer = 0
    Private midHeightCirclePositionZ As Integer = 0
    Private roofCirclePositionZ As Integer = 0

    Private GroundFloorFilePath As String = ""
    Private MidHeightFilePath As String = ""
    Private RoofDeckFilePath As String = ""

    Private visualizationStopwatch As New Stopwatch()

    Private Sub SelectRecordingToVisualize_Click(sender As Object, e As EventArgs) Handles SelectGroundFloorFileButton.Click, SelectMidHeightFileButton.Click, SelectRoofDeckFileButton.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            Dim filePath As String = openFileDialog.FileName
            ' Get only the filename and add it to the appropriate TextBox
            Dim fileName As String = Path.GetFileName(filePath)

            Select Case sender.Name
                Case "SelectGroundFloorFileButton"
                    SelectGroundFloorTextBox.Text = fileName
                    ' Optionally, store the full path if needed
                    GroundFloorFilePath = filePath

                Case "SelectMidHeightFileButton"
                    SelectMidHeightTextBox.Text = fileName
                    ' Optionally, store the full path if needed
                    MidHeightFilePath = filePath

                Case "SelectRoofDeckFileButton"
                    SelectRoofDeckTextBox.Text = fileName
                    ' Optionally, store the full path if needed
                    RoofDeckFilePath = filePath
            End Select
        End If
    End Sub

    Private Sub ERISMainTabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ERISMainTabControl.SelectedIndexChanged
        InitializeCirclePositions()
    End Sub

    Private Sub InitializeCirclePositions()
        'groundCirclePositionX = (FrontViewPictureBox.Width / 2) - 10
        'midHeightCirclePositionX = (FrontViewPictureBox.Width / 2) - 10
        'roofCirclePositionX = (FrontViewPictureBox.Width / 2) - 10

        'groundCirclePositionY = (FrontViewPictureBox.Width / 2) - 10
        'midHeightCirclePositionY = (FrontViewPictureBox.Width / 2) - 10
        'roofCirclePositionY = (FrontViewPictureBox.Width / 2) - 10

        'groundCirclePositionZ = (FrontViewPictureBox.Width / 2) - 10
        'midHeightCirclePositionZ = (FrontViewPictureBox.Width / 2) - 10
        'roofCirclePositionZ = (FrontViewPictureBox.Width / 2) - 10
    End Sub

    Private Sub StartVisualizationButton_Click(sender As Object, e As EventArgs) Handles StartVisualizationButton.Click
        ' Load and process the acceleration data
        'X
        groundXDisplacementData = LoadAndConvertXDisplacementData(GroundFloorFilePath)
        midHeightXDisplacementData = LoadAndConvertXDisplacementData(MidHeightFilePath)
        roofXDisplacementData = LoadAndConvertXDisplacementData(RoofDeckFilePath)
        'Y
        groundYDisplacementData = LoadAndConvertYDisplacementData(GroundFloorFilePath)
        midHeightYDisplacementData = LoadAndConvertYDisplacementData(MidHeightFilePath)
        roofYDisplacementData = LoadAndConvertYDisplacementData(RoofDeckFilePath)
        'Z
        groundZDisplacementData = LoadAndConvertZDisplacementData(GroundFloorFilePath)
        midHeightZDisplacementData = LoadAndConvertZDisplacementData(MidHeightFilePath)
        roofZDisplacementData = LoadAndConvertZDisplacementData(RoofDeckFilePath)

        ' Start the visualization
        VisualizationTimer.Start()
        visualizationStopwatch.Restart() ' Start the stopwatch
        InitializeCirclePositions()
    End Sub

    Private Function LoadAndConvertXDisplacementData(filePath As String) As List(Of Double)
        Dim displacementData As New List(Of Double)
        Dim velocity As Double = 0
        Dim displacement As Double = 0
        Dim deltaTime As Double = 0.01 ' Assuming a fixed time interval (0.01 seconds)

        ' Read the acceleration data from the file
        Dim lines() As String = File.ReadAllLines(filePath)

        ' Skip the first 6 lines (metadata)
        For i As Integer = 6 To lines.Length - 2
            Dim parts() As String = lines(i).Split(vbTab)
            Dim nextParts() As String = lines(i + 1).Split(vbTab)

            Dim accelX As Double = Double.Parse(parts(1))
            Dim nextAccelX As Double = Double.Parse(nextParts(1))

            ' Calculate the change in velocity using trapezoidal rule
            Dim deltaV As Double = 0.5 * (accelX + nextAccelX) * deltaTime

            ' Update velocity
            velocity += deltaV

            ' Update displacement
            displacement += velocity * deltaTime

            Dim scalingFactor As Integer = 1000

            ' Store displacement value
            displacementData.Add(displacement * scalingFactor)
        Next

        Return displacementData
    End Function
    Private Function LoadAndConvertYDisplacementData(filePath As String) As List(Of Double)
        Dim displacementData As New List(Of Double)
        Dim velocity As Double = 0
        Dim displacement As Double = 0
        Dim deltaTime As Double = 0.01 ' Assuming a fixed time interval (0.01 seconds)

        ' Read the acceleration data from the file
        Dim lines() As String = File.ReadAllLines(filePath)

        ' Skip the first 6 lines (metadata)
        For i As Integer = 6 To lines.Length - 2
            Dim parts() As String = lines(i).Split(vbTab)
            Dim nextParts() As String = lines(i + 1).Split(vbTab)

            Dim accelY As Double = Double.Parse(parts(2))
            Dim nextAccelY As Double = Double.Parse(nextParts(2))

            ' Calculate the change in velocity using trapezoidal rule
            Dim deltaV As Double = 0.5 * (accelY + nextAccelY) * deltaTime

            ' Update velocity
            velocity += deltaV

            ' Update displacement
            displacement += velocity * deltaTime

            Dim scalingFactor As Integer = 1000

            ' Store displacement value
            displacementData.Add(displacement * scalingFactor)
        Next

        Return displacementData
    End Function
    Private Function LoadAndConvertZDisplacementData(filePath As String) As List(Of Double)
        Dim displacementData As New List(Of Double)
        Dim velocity As Double = 0
        Dim displacement As Double = 0
        Dim deltaTime As Double = 0.01 ' Assuming a fixed time interval (0.01 seconds)

        ' Read the acceleration data from the file
        Dim lines() As String = File.ReadAllLines(filePath)

        ' Skip the first 6 lines (metadata)
        For i As Integer = 6 To lines.Length - 2
            Dim parts() As String = lines(i).Split(vbTab)
            Dim nextParts() As String = lines(i + 1).Split(vbTab)

            Dim accelZ As Double = Double.Parse(parts(3))
            Dim nextAccelZ As Double = Double.Parse(nextParts(3))

            ' Calculate the change in velocity using trapezoidal rule
            Dim deltaV As Double = 0.5 * (accelZ + nextAccelZ) * deltaTime

            ' Update velocity
            velocity += deltaV

            ' Update displacement
            displacement += velocity * deltaTime

            Dim scalingFactor As Integer = 1000

            ' Store displacement value
            displacementData.Add(displacement * scalingFactor)
        Next

        Return displacementData
        Return displacementData
    End Function

    Private Sub VisualizationTimer_Tick(sender As Object, e As EventArgs) Handles VisualizationTimer.Tick
        ' Update the circle positions based on the displacement data
        If groundCurrentIndex < groundXDisplacementData.Count - 1 AndAlso
           midHeightCurrentIndex < midHeightXDisplacementData.Count - 1 AndAlso
           roofCurrentIndex < roofXDisplacementData.Count - 1 Then

            groundCirclePositionX = (groundXDisplacementData(groundCurrentIndex))
            midHeightCirclePositionX = (midHeightXDisplacementData(midHeightCurrentIndex))
            roofCirclePositionX = (roofXDisplacementData(roofCurrentIndex))

            groundCirclePositionY = (groundYDisplacementData(groundCurrentIndex))
            midHeightCirclePositionY = (midHeightYDisplacementData(midHeightCurrentIndex))
            roofCirclePositionY = (roofYDisplacementData(roofCurrentIndex))

            groundCirclePositionZ = (groundZDisplacementData(groundCurrentIndex))
            midHeightCirclePositionZ = (midHeightZDisplacementData(midHeightCurrentIndex))
            roofCirclePositionZ = (roofZDisplacementData(roofCurrentIndex))

            groundCurrentIndex += 1
            midHeightCurrentIndex += 1
            roofCurrentIndex += 1

            FrontViewPictureBox.Invalidate() ' Refresh the drawing
            SideViewPictureBox.Invalidate() ' Refresh the drawing
            PlanViewPictureBox.Invalidate() ' Refresh the drawing
        Else
            VisualizationTimer.Stop() ' Stop the timer when data is exhausted
            groundCurrentIndex = 0
            midHeightCurrentIndex = 0
            roofCurrentIndex = 0

            ' Stop the stopwatch and display the elapsed time
            visualizationStopwatch.Stop()
            Dim elapsedTime As Double = visualizationStopwatch.Elapsed.TotalSeconds
            MessageBox.Show("Visualization completed in " & elapsedTime.ToString("F2") & " seconds.")
        End If
    End Sub

    Private Sub FrontViewPictureBox_Paint(sender As Object, e As PaintEventArgs) Handles FrontViewPictureBox.Paint
        ' Draw the circles at the current x positions
        Dim g As Graphics = e.Graphics
        Dim circleRadius As Integer = 10

        Dim x1 As Integer = (FrontViewPictureBox.Width / 2) - circleRadius + groundCirclePositionX ' Ground Floor X
        Dim x2 As Integer = (FrontViewPictureBox.Width / 2) - circleRadius + midHeightCirclePositionX ' Mid Height X
        Dim x3 As Integer = (FrontViewPictureBox.Width / 2) - circleRadius + roofCirclePositionX ' Roof Deck X

        Dim z1 As Integer = FrontViewPictureBox.Height * (3 / 4) - circleRadius + groundCirclePositionZ ' Ground Floor Z
        Dim z2 As Integer = FrontViewPictureBox.Height * (1 / 2) - circleRadius + midHeightCirclePositionZ ' Mid Height Z
        Dim z3 As Integer = FrontViewPictureBox.Height * (1 / 4) - circleRadius + roofCirclePositionZ ' Roof Deck Z

        g.Clear(Color.White) ' Clear the previous drawing

        ' Draw the circles
        g.FillEllipse(Brushes.Red, x3, z3, circleRadius * 2, circleRadius * 2) ' Roof Deck Movement
        g.FillEllipse(Brushes.Orange, x2, z2, circleRadius * 2, circleRadius * 2) ' Mid Height Movement
        g.FillEllipse(Brushes.Blue, x1, z1, circleRadius * 2, circleRadius * 2) ' Ground Floor Movement

    End Sub
    Private Sub SideViewPictureBox_Paint(sender As Object, e As PaintEventArgs) Handles SideViewPictureBox.Paint
        ' Draw the circles at the current x positions
        Dim g As Graphics = e.Graphics
        Dim circleRadius As Integer = 10

        Dim y1 As Integer = (SideViewPictureBox.Width / 2) - circleRadius + groundCirclePositionY ' Ground Floor Y
        Dim y2 As Integer = (SideViewPictureBox.Width / 2) - circleRadius + midHeightCirclePositionY ' Mid Height Y
        Dim y3 As Integer = (SideViewPictureBox.Width / 2) - circleRadius + roofCirclePositionY ' Roof Deck Y

        Dim z1 As Integer = SideViewPictureBox.Height * (3 / 4) - circleRadius + groundCirclePositionZ ' Ground Floor Z
        Dim z2 As Integer = SideViewPictureBox.Height * (1 / 2) - circleRadius + midHeightCirclePositionZ ' Mid Height Z
        Dim z3 As Integer = SideViewPictureBox.Height * (1 / 4) - circleRadius + roofCirclePositionZ ' Roof Deck Z

        g.Clear(Color.White) ' Clear the previous drawing

        ' Draw the circles
        g.FillEllipse(Brushes.Red, y3, z3, circleRadius * 2, circleRadius * 2) ' Roof Deck Movement
        g.FillEllipse(Brushes.Orange, y2, z2, circleRadius * 2, circleRadius * 2) ' Mid Height Movement
        g.FillEllipse(Brushes.Blue, y1, z1, circleRadius * 2, circleRadius * 2) ' Ground Floor Movement

    End Sub
    Private Sub PlanViewPictureBox_Paint(sender As Object, e As PaintEventArgs) Handles PlanViewPictureBox.Paint
        ' Draw the circles at the current x positions
        Dim g As Graphics = e.Graphics
        Dim circleRadius As Integer = 10

        Dim x1 As Integer = (PlanViewPictureBox.Width / 2) - circleRadius + groundCirclePositionX ' Ground Floor X
        Dim x2 As Integer = (PlanViewPictureBox.Width / 2) - circleRadius + midHeightCirclePositionX ' Mid Height X
        Dim x3 As Integer = (PlanViewPictureBox.Width / 2) - circleRadius + roofCirclePositionX ' Roof Deck X

        Dim y1 As Integer = PlanViewPictureBox.Height * (1 / 2) - circleRadius + groundCirclePositionY ' Ground Floor Y
        Dim y2 As Integer = PlanViewPictureBox.Height * (1 / 2) - circleRadius + midHeightCirclePositionY ' Mid Height Y
        Dim y3 As Integer = PlanViewPictureBox.Height * (1 / 2) - circleRadius + roofCirclePositionY ' Roof Deck Y

        g.Clear(Color.White) ' Clear the previous drawing

        ' Draw the circles
        g.FillEllipse(Brushes.Blue, x1, y1, circleRadius * 2, circleRadius * 2) ' Ground Floor Movement
        g.FillEllipse(Brushes.Orange, x2, y2, circleRadius * 2, circleRadius * 2) ' Mid Height Movement
        g.FillEllipse(Brushes.Red, x3, y3, circleRadius * 2, circleRadius * 2) ' Roof Deck Movement


    End Sub
#End Region


End Class


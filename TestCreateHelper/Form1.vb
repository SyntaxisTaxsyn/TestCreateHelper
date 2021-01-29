Imports System.IO
Imports System.Reflection
Imports System.Linq
Public Class Form1

    Public STR_PathFile As String = ""
    Public ValuesList As List(Of ValuePair)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' Display File Picker for HMI XML export file
        Me.OpenFileDialog1.Filter = "XML FILES (.XML)|*.XML" ' set file filter
        Me.OpenFileDialog1.Title = "Select the HMI XML Display Export"
        Me.OpenFileDialog1.ShowDialog(Me) 'display file picker
        STR_PathFile = OpenFileDialog1.FileName.ToString 'assign file path to variable

        If STR_PathFile Like "*OpenFileDialog*" Then
            MsgBox("No file picked")
            Application.Exit()
        End If

        Me.TextBox1.Text = STR_PathFile

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' workflow
        ' open the file
        ' locate the xml object (the file will have only one as i will isolate it prior)
        ' find all the parameters in the header object
        ' create a csv from the parameters 
        ' create 2 pages left and right that have the parameters changed like the csv

        Dim tstr(100) As String
        Dim linecount As Integer = 0
        Dim MainObjectList As New ParamList

        If STR_PathFile = "" Then
            MsgBox("Unable to compare, file name missing")
            Exit Sub
        End If

        If Not File.Exists(STR_PathFile) Then
            MsgBox("Error - Unable to compare, Left file does not exist")
            Exit Sub

        End If

        Using reader As StreamReader = New StreamReader(STR_PathFile)
            Do
                tstr(linecount) = reader.ReadLine
                linecount += 1

            Loop Until reader.EndOfStream
        End Using

        If linecount = 1 Then
            ' this is a single line object definition
            MainObjectList = GetObjParams(tstr(0))
            'MsgBox("")
        Else
            ' this object has sub objects
            tstr = tstr

            ' For each line find out the type and do the appropriate action
            ' The first line is always the header object so we can process this as normal
            MainObjectList = GetObjParams(tstr(0))
            MainObjectList.sList = New SubParamList ' initialise the sub object parameter list as its required for this object type
            Dim type As String = ""
            Dim connectionsPresent As Boolean = False
            Dim connectionLRange As Integer = 0
            Dim connectionRRange As Integer = 0
            ' Add state handling variables
            Dim statesPresent As Boolean = False

            Dim stateLRange As Integer = 0
            Dim stateRRange As Integer = 0
            ' Add threshold handling variables
            Dim thresholdsPresent As Boolean = False
            Dim thresholdLRange As Integer = 0
            Dim thresholdRRange As Integer = 0
            ' Process the remaining lines
            For a = 1 To linecount - 1
                Dim typeIsKnown As Boolean = False
                If InStr(tstr(a), "<caption") Then
                    type = TypeConstants.caption
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "<imageSettings") Then
                    type = TypeConstants.imageSettings
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "<connections>") Then
                    type = TypeConstants.connections
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "<connection name") Then
                    type = TypeConstants.connection_name
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "</") > 0 Then
                    type = TypeConstants.ClosingTag
                    typeIsKnown = True
                End If
                If tstr(a) = "" Then
                    ' an empty line just needs ignored
                    type = ""
                    typeIsKnown = True
                End If
                ' Add new state type handling here
                If InStr(tstr(a), "<states>") Then
                    type = TypeConstants.states
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "<state ") Then
                    type = TypeConstants.state
                    typeIsKnown = True
                End If
                ' Add new threshold type handling here
                If InStr(tstr(a), "<threshold") Then
                    type = TypeConstants.Threshold
                    typeIsKnown = True
                End If
                ' Add new activeX data type handling here
                If InStr(tstr(a), "<data") Then
                    type = TypeConstants.Data
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "<animations>") Then
                    type = TypeConstants.Animation
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "<animate") Then
                    type = TypeConstants.Animate
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "<color") Then
                    type = TypeConstants.Animate
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "readFromTagExpressionRange") Then
                    type = TypeConstants.readFromTagExpressionRange
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "constantExpressionRange") Then
                    type = TypeConstants.constantExpressionRange
                    typeIsKnown = True
                End If
                If InStr(tstr(a), "defaultExpressionRange") Then
                    type = TypeConstants.defaultExpressionRange
                    typeIsKnown = True
                End If
                If Not typeIsKnown Then
                    ' This code will prompt us for any rework required going forward as we encounter new sub types
                    Throw New Exception("Unknown Type Detected, Requires new code to handle" & vbCrLf &
                                        "new type contained in - " & tstr(a))
                End If
                Select Case type
                    Case TypeConstants.caption
                        ' add a new subobject to the main object and parse the complete caption line as normal
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                        'MsgBox("")
                    Case TypeConstants.imageSettings
                        ' add a new subobject to the main object and parse the complete caption line as normal
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case TypeConstants.connections
                        ' each connection is its own subobject, this needs some special processing
                        ' find the range of connection entries to process
                        If connectionsPresent = False Then
                            ' determine range first
                            connectionLRange = a + 1 ' miss the current line as it only contains "<connections>" to define subobject group
                            ' Find the number of connections present
                            For c = connectionLRange To linecount - 1
                                If InStr(tstr(c), "</connections>") > 0 Then
                                    connectionRRange = c - 1
                                    connectionsPresent = True
                                End If
                            Next
                        Else
                        End If
                    Case TypeConstants.connection_name
                        If connectionsPresent Then
                            If a >= connectionLRange And a <= connectionRRange Then
                                Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                                MainObjectList.sList.lSubParamList.Add(newSubObject)
                            End If
                        End If

                    Case TypeConstants.state ' Add new case handling for state types

                        If a >= stateLRange And a <= stateRRange Then
                            Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                            MainObjectList.sList.lSubParamList.Add(newSubObject)
                        End If

                    Case TypeConstants.states
                        ' each connection is its own subobject, this needs some special processing
                        ' find the range of connection entries to process
                        If statesPresent = False Then
                            ' determine range first
                            stateLRange = a + 1 ' miss the current line as it only contains "<connections>" to define subobject group
                            ' Find the number of connections present
                            For c = stateLRange To linecount - 1
                                If InStr(tstr(c), "</states>") > 0 Then
                                    stateRRange = c - 1
                                    statesPresent = True
                                End If
                            Next
                        Else
                        End If

                    Case TypeConstants.Animate
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case TypeConstants.Animation
                        ' Do nothing in here, this just avoids throwinng an exception
                    Case TypeConstants.Threshold
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case "</"
                    Case ""
                    Case TypeConstants.Data
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case TypeConstants.Color
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case TypeConstants.readFromTagExpressionRange
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case TypeConstants.constantExpressionRange
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case TypeConstants.defaultExpressionRange
                        Dim newSubObject As subparam = GetSubObjParams(tstr(a))
                        MainObjectList.sList.lSubParamList.Add(newSubObject)
                    Case Else
                        Throw New Exception("Unhandled type detected")

                End Select
            Next
            'MsgBox("")
        End If

        ' Load the dictionary CSV to get all the name/value pairs
        Dim filelist As List(Of String) = New List(Of String)
        Using reader As StreamReader = New StreamReader(GetPathToDefFile("TestDefinitions"))
            Do
                filelist.Add(reader.ReadLine)
            Loop Until reader.EndOfStream
        End Using

        ' Convert the CSV file into a list value pair type objects
        Dim ValuePairList As List(Of ValuePair) = New List(Of ValuePair)
        For items = 0 To filelist.Count - 1
            Dim tr()
            tr = Split(filelist.Item(items), ",")
            Dim DuplicateFound As Boolean = False
            For Each itm In ValuePairList
                ' Filter the csv file for duplicates
                If DuplicateFound = False Then
                    If tr(0) = itm.oClass Then
                        If tr(1) = itm.name Then
                            DuplicateFound = True
                            Exit For ' no point in continuing this for each loop as this item is a match
                        End If
                    End If
                End If
            Next
            If DuplicateFound = False Then
                ' you can add this value pair to the list as it has no prior duplicates found
                ValuePairList.Add(New ValuePair(tr(0), tr(1), ReturnFormattedValues(tr(2)), ReturnFormattedValues(tr(3))))
            End If
        Next

        ' Now cross check the test object data against the test definitions and flag up any items that dont exist so they can be added manually

        Dim ExceptionList As List(Of String) = New List(Of String)
        Dim bFoundMatch As Boolean

        ' first loop through the main parameter list then check the sub parameter lists as well
        For Each itm As Param In MainObjectList.pList
            bFoundMatch = False
            For Each oValPair As ValuePair In ValuePairList
                If oValPair.name = itm.sProperty Then
                    bFoundMatch = True
                End If
            Next
            If Not bFoundMatch Then
                Select Case itm.sProperty
                    ' Filter out exception cases with an empty select case statement
                    ' The final case else catches unfiltered properties and throws them as esceptions
                    Case "name" ' none of these cases do anything, they are merely for exception filtering
                        'Case "wallpaper"
                        'Case "isReferenceObject'"
                    Case Else
                        ExceptionList.Add(itm.sProperty)
                End Select
            End If
        Next

        If ExceptionList.Count > 0 Then
            Outputreport(ExceptionList)
            Throw New Exception("Missing test definitions detected, observe output file and correct manually")
        End If

        ' we got this far, check for declared sub parameter types and deal with them accordingly
        If MainObjectList.sList IsNot Nothing Then
            For Each sublist As subparam In MainObjectList.sList.lSubParamList
                For Each itm As Param In sublist.subParList
                    bFoundMatch = False
                    For Each oValPair As ValuePair In ValuePairList
                        If oValPair.name = itm.sProperty Then
                            If oValPair.oClass = TypeConstants.Animate Then
                                ' Handle special case for animation type class objects
                                If sublist.type Like TypeConstants.Animate & "*" Then
                                    bFoundMatch = True
                                End If
                            Else
                                ' do the class matching in the normal way
                                If oValPair.oClass = sublist.type Then ' Added this selector in to ensure that value pairs get matched against classes too
                                    bFoundMatch = True
                                End If
                            End If

                        End If
                    Next
                    If Not bFoundMatch Then
                        Select Case itm.sProperty
                            ' Filter out exception cases with an empty select case statement
                            ' The final case else catches unfiltered properties and throws them as esceptions
                            Case "name" ' none of these cases do anything, they are merely for exception filtering
                            Case Else

                                ExceptionList.Add(sublist.type & "-" & itm.sProperty)
                        End Select
                    End If
                Next
            Next
        End If

        If ExceptionList.Count > 1 Then
            Outputreport(ExceptionList)
            Throw New Exception("Missing test definitions detected, observe output file and correct manually")
        End If

        ' if we get this far with no exceptions now its time to start generating the files
        ' read in the header and footer files for the generation process
        Dim HeaderList As List(Of String) = New List(Of String)
        Dim FooterList As List(Of String) = New List(Of String)

        HeaderList = ReadFile(HeaderList, GetPathToLocalFile("Test Definitions", "Header.xml"))
        FooterList = ReadFile(FooterList, GetPathToLocalFile("Test Definitions", "Footer.xml"))

#Region "Main_List_Generation"

        ' Generate file content for the main parameter list
        Dim MainFileListLeft As List(Of String) = New List(Of String)
        Dim MainFileListRight As List(Of String) = New List(Of String)
        Dim MainFileCSV As List(Of String) = New List(Of String)
        Dim TestCount As Integer = 1
        Dim OType As ECloseType
        Dim FirstConnectionFound As Boolean = False
        Dim FirstStateFound As Boolean = False
        Dim FirstAnimationFound As Boolean = False
        Dim FirstColorFound As Boolean = False
        Dim FirstCaptionFound As Boolean = False
        Dim CaptionIndentLevel As Integer = 1
        Dim StatesClosed As Boolean = False
        Dim AnimationsClosed As Boolean = False
        Dim ColorsClosed As Boolean = False
        Dim Type_ImageSettings_Exists As Boolean = False
        Dim Type_Caption_Exists As Boolean = False
        Dim Type_Threshold_Exists As Boolean = False
        Dim Type_Connection_Exists As Boolean = False
        Dim Type_State_Exists As Boolean = False
        Dim Type_ActiveXData_Exists As Boolean
        Dim Type_Animation_Exists As Boolean = False
        Dim Type_ExpressionRange_Exists As Boolean = True
        Dim Type_Color_Exists As Boolean = False
        Dim StateCount As Integer = 0
        Dim CaptionCount As Integer = 0
        Dim ImageCount As Integer = 0
        Dim ThresholdCount As Integer = 0
        Dim DataCount As Integer
        Dim TagToClose As String = ""

        If MainObjectList.sList IsNot Nothing Then
            OType = ECloseType.Complex
        Else
            OType = ECloseType.Simple
        End If

        ' Initialise the left and right file lists with the header content
        For Each itm As String In HeaderList
            MainFileListLeft.Add(itm)
            MainFileListRight.Add(itm)
        Next

        ' Initialise the CSV test definition file list
        MainFileCSV.Add("Test number,Property,Left,Right")


        For Each itm As Param In MainObjectList.pList
            If Not itm.sProperty = "name" Then

                MainFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, itm, EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                MainFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, itm, EditCase.Right, 0, TestCount, ValuePairList, "", OType))
                MainFileCSV.Add(CreateTestCaseByTestNumber(itm, ValuePairList, "", TestCount))
                If MainObjectList IsNot Nothing Then
                    If MainObjectList.sList IsNot Nothing Then
                        ' Enumerate through the sub object children and create those entries for this object instance
                        For Each subp As subparam In MainObjectList.sList.lSubParamList
                            Select Case subp.type
                                Case TypeConstants.Data
                                    MainFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.data,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.data,
                                                                                     ECloseType.Simple))
                                Case TypeConstants.Threshold
                                    MainFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.threshold,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.threshold,
                                                                                     ECloseType.Simple))
                                Case TypeConstants.caption
                                    Select Case FirstStateFound
                                        Case True
                                            CaptionIndentLevel = 3
                                        Case False
                                            CaptionIndentLevel = 1
                                    End Select
                                    MainFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     CaptionIndentLevel,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.caption,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     CaptionIndentLevel,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.caption,
                                                                                     ECloseType.Simple))
                                Case TypeConstants.imageSettings
                                    MainFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.image,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.image,
                                                                                     ECloseType.Simple))
                                Case TypeConstants.connection
                                    If FirstAnimationFound Then
                                        If Not AnimationsClosed Then
                                            'If FirstColorFound Then
                                            '    If Not ColorsClosed Then
                                            '        ' close off the animatecolor block
                                            '        MainFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                            '        MainFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                            '        ColorsClosed = True
                                            '    End If
                                            'End If
                                            If Not TagToClose = "" Then
                                                MainFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                MainFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                TagToClose = ""
                                            End If
                                            ' Add lines here to close off the animation objects
                                            MainFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                            MainFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                            AnimationsClosed = True
                                        End If
                                    End If
                                    If FirstStateFound Then
                                        If Not StatesClosed Then
                                            ' Close off the previous state before starting a connection block
                                            MainFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                            MainFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            MainFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                            MainFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                            StatesClosed = True
                                        End If

                                    End If
                                    If FirstConnectionFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        MainFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                        MainFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                        FirstConnectionFound = True
                                    End If
                                    MainFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.caption,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.caption,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.state
                                    If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                        ' if this is a subsequent state found after the first then close off the previous state
                                        MainFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        MainFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                    End If
                                    If FirstStateFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        MainFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                        MainFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                        FirstStateFound = True
                                    End If
                                    MainFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.state,
                                                                                     ECloseType.Complex))
                                    MainFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.state,
                                                                                    ECloseType.Complex))
                                Case TypeConstants.Color
                                    If FirstColorFound = False Then
                                        FirstColorFound = True
                                    End If
                                    MainFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.color,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.color,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.readFromTagExpressionRange
                                    MainFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.readfromtagexpressionrange,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.readfromtagexpressionrange,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.constantExpressionRange
                                    MainFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.constantexpressionrange,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.constantexpressionrange,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.defaultExpressionRange
                                    MainFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.defaultexpressionrange,
                                                                                     ECloseType.Simple))
                                    MainFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.defaultexpressionrange,
                                                                                    ECloseType.Simple))
                                Case Else
                                    ' Handle the animation types in here as we will be bundling all types into the same class hence a like
                                    ' comparison operator is used to check the object class type
                                    If subp.type Like TypeConstants.Animate & "*" Then
                                        If FirstAnimationFound = False Then
                                            ' Add an additional line here for the connection xml configuration on the first time only
                                            MainFileListLeft.Add(AddWhiteSpace(1, "<animations>"))
                                            MainFileListRight.Add(AddWhiteSpace(1, "<animations>"))
                                            FirstAnimationFound = True
                                        End If
                                        Dim SelectEcloseType As ECloseType = GetAnimationEcloseType(subp.type)
                                        If SelectEcloseType = ECloseType.Complex Then
                                            ' Store the name of the animation tag so we can close it later
                                            ' Also check if the tag to store has changed so we can close previous tags
                                            If Not TagToClose = "" Then
                                                ' case when closing tag already exists
                                                If TagToClose = subp.type Then
                                                    ' Do nothing because this is the same type as before, it will be closed at the end
                                                    ' Of the main loop
                                                Else
                                                    ' It is a different type, close out the old one and start a new tag
                                                    MainFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    MainFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    TagToClose = subp.type
                                                    MainFileListLeft.Add(AddWhiteSpace(1, "<" & TagToClose & ">"))
                                                    MainFileListRight.Add(AddWhiteSpace(1, "<" & TagToClose & ">"))
                                                End If
                                            Else
                                                ' tagtoclose not set yet so update it with current type
                                                TagToClose = subp.type
                                            End If
                                        Else
                                            ' Upon encountering a simple type check if a previous complex type needs closed first
                                            If Not TagToClose = "" Then
                                                ' a previous tag needs closed first
                                                MainFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                MainFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                TagToClose = "" ' this lets us know on the next loop that nothing requires closing
                                            Else
                                                ' Do nothing
                                                ' no tags opened to be closed and this is a simple type so just deal with it normally
                                            End If
                                        End If
                                        MainFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.animate,
                                                                                         SelectEcloseType))
                                        MainFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.animate,
                                                                                         SelectEcloseType))
                                    Else
                                        Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End If

                            End Select
                        Next
                        If FirstAnimationFound Then
                            If Not AnimationsClosed Then
                                'If FirstColorFound Then
                                '    If Not ColorsClosed Then
                                '        ' close off the animatecolor block
                                '        MainFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                '        MainFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                '        ColorsClosed = True
                                '    End If
                                'End If
                                If Not TagToClose = "" Then
                                    MainFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                    MainFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                    TagToClose = ""
                                End If
                                ' Add lines here to close off the animation objects
                                MainFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                MainFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                AnimationsClosed = True
                            End If
                        End If

                        ' Reset first found and closed tag monitors here
                        FirstAnimationFound = False
                        AnimationsClosed = False
                        ColorsClosed = False
                        FirstColorFound = False

                        If FirstStateFound Then
                            If Not StatesClosed Then
                                ' handle the case when no connection block is present and the state blocks need closed
                                MainFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                MainFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                MainFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                MainFileListRight.Add(AddWhiteSpace(1, "</states>"))
                            End If
                            StatesClosed = False
                            FirstStateFound = False
                        End If
                        If FirstConnectionFound Then
                            ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                            ' We do this at the end of the sub group iteration because we know by observation of the ME software
                            ' XML object creation that connections always go at the end
                            MainFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                            MainFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                            FirstConnectionFound = False
                        End If
                        ' Close off this XML object
                        If OType = ECloseType.Complex Then
                            ' Requires complex object closure
                            MainFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                            MainFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                        End If
                    End If
                Else
                    ' this is a simple type so no other work required here, the object is already closed off

                End If

            End If
            TestCount += 1
        Next

        ' Close off the files with the footer
        For Each itm As String In FooterList
            MainFileListLeft.Add(itm)
            MainFileListRight.Add(itm)
        Next

        'Format output file contents prior to writing
        FormatXMLFile(MainFileListLeft)
        FormatXMLFile(MainFileListRight)

        Dim FnameVar As String = InputBox("Enter Output file name", "")

        WriteOutputFile(MainFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_main.xml"))
        WriteOutputFile(MainFileListRight, GetPathToLocalFile("Output", FnameVar & "R_main.xml"))
        WriteOutputFile(MainFileCSV, GetPathToLocalFile("Output", FnameVar & "main.csv"))
#End Region

#Region "Image List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type = TypeConstants.imageSettings Then
                    Type_ImageSettings_Exists = True
                End If
            Next
        End If

        If Type_ImageSettings_Exists Then
            ' Generate file content for the main parameter list
            Dim ImageFileListLeft As List(Of String) = New List(Of String)
            Dim ImageFileListRight As List(Of String) = New List(Of String)
            Dim ImageFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False
            FirstCaptionFound = False ' Added to ensure only the first caption type gets processed when dealing with mutlistate objects
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            ImageCount = 0
            Dim ImageMask(10) As Boolean
            Dim StateInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.state)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                ImageFileListLeft.Add(itm)
                ImageFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            ImageFileCSV.Add("Test number,Property,Left,Right")

            ' Set up caption mask
            Select Case True
                Case StateInstCount > 3
                    ImageMask(0) = True
                    ImageMask(1) = True
                    ImageMask(3) = True
                Case StateInstCount > 2
                    ImageMask(0) = True
                    ImageMask(1) = True
                    ImageMask(2) = True
                Case StateInstCount = 1
                    ImageMask(0) = True
                Case StateInstCount = 2
                    ImageMask(0) = True
                    ImageMask(1) = True
                Case StateInstCount = 0
                    ImageMask(0) = True
                Case Else
                    Throw New Exception("Whoops, it appears you didnt think of everything")
            End Select

            ' Loop through the test generation process for as many caption test masks are active
            For Imask = 0 To 9
                If ImageMask(Imask) Then
                    For Each sublist As subparam In MainObjectList.sList.lSubParamList
                        If sublist.type = TypeConstants.imageSettings Then
                            ' This object has a caption sub object type so generate an output file for it
                            For Each sparam As Param In sublist.subParList
                                ' Generate the main objects data with only left cases
                                ImageFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                ImageFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                'ImageFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.caption, TestCount))

                                For Each subp As subparam In MainObjectList.sList.lSubParamList
                                    Select Case subp.type
                                        Case TypeConstants.Data
                                            ImageFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                            ImageFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.Threshold
                                            ImageFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                            ImageFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.caption
                                            ImageFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.caption,
                                                                                         ECloseType.Simple))
                                            ImageFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.caption,
                                                                                        ECloseType.Simple))

                                        Case TypeConstants.imageSettings

                                            If ImageCount = Imask Then ' Select the image instance count to modify based on the mask
                                                ' Add test case for this image only
                                                Dim Addstr As String = DetermineAddStrByCase(MainObjectList, (StateCount - 1)) ' State number - 1 here because each state clause starts before the image clause hence the count will +1
                                                ImageFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.image, TestCount, Addstr))
                                                ' Only substitute params in the first image object
                                                ImageFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                              sparam,
                                                                                              EditCase.Left,
                                                                                              CaptionIndentLevel,
                                                                                              TestCount,
                                                                                              ValuePairList,
                                                                                              ObjectTestClass.image,
                                                                                              ECloseType.Simple))

                                                ImageFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Right,
                                                                                                  CaptionIndentLevel,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.image,
                                                                                                  ECloseType.Simple))
                                            Else
                                                ' Add subsequent captions with left case (default) parameters only
                                                ImageFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                                ImageFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.image,
                                                                                                 ECloseType.Simple))
                                            End If
                                            ImageCount += 1

                                        Case TypeConstants.connection
                                            If FirstStateFound Then
                                                ' Close off the previous state before starting a connection block
                                                ImageFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                ImageFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                                ImageFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                                ImageFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                                StatesClosed = True
                                            End If
                                            If FirstConnectionFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                ImageFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                                ImageFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                                FirstConnectionFound = True
                                            End If
                                            ImageFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                            ImageFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                        Case TypeConstants.state
                                            If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                                ' if this is a subsequent state found after the first then close off the previous state
                                                ImageFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                ImageFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            End If
                                            If FirstStateFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                ImageFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                                ImageFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                                FirstStateFound = True
                                            End If
                                            ImageFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.state,
                                                                                         ECloseType.Complex))
                                            ImageFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.state,
                                                                                        ECloseType.Complex))
                                            StateCount += 1
                                        Case Else
                                            Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End Select

                                Next
                                If FirstStateFound Then
                                    If Not StatesClosed Then
                                        ' handle the case when no connection block is present and the state blocks need closed
                                        ImageFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        ImageFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                        ImageFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                        ImageFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                    End If
                                    StatesClosed = False
                                    FirstStateFound = False
                                    StateCount = 0
                                End If
                                If FirstConnectionFound Then
                                    ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                                    ' We do this at the end of the sub group iteration because we know by observation of the ME software
                                    ' XML object creation that connections always go at the end
                                    ImageFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                                    ImageFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                                    FirstConnectionFound = False
                                End If
                                ImageCount = 0

                                ' Close off this XML object
                                If OType = ECloseType.Complex Then
                                    ' Requires complex object closure
                                    ImageFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                    ImageFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                End If
                                TestCount += 1
                                FirstCaptionFound = False
                            Next
                            Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                        End If
                    Next
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                ImageFileListLeft.Add(itm)
                ImageFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(ImageFileListLeft)
            FormatXMLFile(ImageFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(ImageFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_image.xml"))
            WriteOutputFile(ImageFileListRight, GetPathToLocalFile("Output", FnameVar & "R_image.xml"))
            WriteOutputFile(ImageFileCSV, GetPathToLocalFile("Output", FnameVar & "image.csv"))


            MsgBox("")

        End If

#End Region

#Region "Caption List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type = TypeConstants.caption Then
                    Type_Caption_Exists = True
                End If
            Next
        End If


        If Type_Caption_Exists Then
            ' Generate file content for the main parameter list
            Dim CaptionFileListLeft As List(Of String) = New List(Of String)
            Dim CaptionFileListRight As List(Of String) = New List(Of String)
            Dim CaptionFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False
            FirstCaptionFound = False ' Added to ensure only the first caption type gets processed when dealing with mutlistate objects
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            CaptionCount = 0
            Dim CaptionMask(10) As Boolean
            Dim StateInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.state)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                CaptionFileListLeft.Add(itm)
                CaptionFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            CaptionFileCSV.Add("Test number,Property,Left,Right")

            'For Each sublist As subparam In MainObjectList.sList.lSubParamList.Where _
            '    (Function(x) x.type = TypeConstants.caption)
            '    ' Linq exp selects the objects by caption, find the first caption, then run through the entire
            '    ' Object structure again generating the page content
            '    For Each sparam As Param In sublist.subParList

            '    Next
            '    Exit For ' this is probably messy but i cant think of another better way to select the first caption object 
            '    ' and then exit without processing the rest
            'Next

            ' Set up caption mask
            Select Case True
                Case StateInstCount > 3
                    CaptionMask(0) = True
                    CaptionMask(1) = True
                    CaptionMask(3) = True
                Case StateInstCount > 2
                    CaptionMask(0) = True
                    CaptionMask(1) = True
                    CaptionMask(2) = True
                Case StateInstCount = 1
                    CaptionMask(0) = True
                Case StateInstCount = 2
                    CaptionMask(0) = True
                    CaptionMask(1) = True
                Case StateInstCount = 0
                    CaptionMask(0) = True
                Case Else
                    Throw New Exception("Whoops, it appears you didnt think of everything")
            End Select

            ' Loop through the test generation process for as many caption test masks are active
            For Cmask = 0 To 9
                If CaptionMask(Cmask) Then
                    For Each sublist As subparam In MainObjectList.sList.lSubParamList
                        If sublist.type = TypeConstants.caption Then
                            ' This object has a caption sub object type so generate an output file for it
                            For Each sparam As Param In sublist.subParList
                                ' Generate the main objects data with only left cases
                                CaptionFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                CaptionFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                'CaptionFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.caption, TestCount))

                                For Each subp As subparam In MainObjectList.sList.lSubParamList
                                    Select Case subp.type
                                        Case TypeConstants.Data
                                            CaptionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                            CaptionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.Threshold
                                            CaptionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                            CaptionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.caption
                                            If CaptionCount = Cmask Then ' Select the caption instance count to modify based on the mask
                                                ' Add test case for this caption only
                                                Dim Addstr As String = DetermineAddStrByCase(MainObjectList, (StateCount - 1)) ' State number - 1 here because each state clause starts before the caption clause hence the count will +1
                                                CaptionFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.caption, TestCount, Addstr))
                                                ' Only substitute params in the first caption object
                                                CaptionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                              sparam,
                                                                                              EditCase.Left,
                                                                                              CaptionIndentLevel,
                                                                                              TestCount,
                                                                                              ValuePairList,
                                                                                              ObjectTestClass.caption,
                                                                                              ECloseType.Simple))

                                                CaptionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Right,
                                                                                                  CaptionIndentLevel,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.caption,
                                                                                                  ECloseType.Simple))
                                                FirstCaptionFound = True
                                            Else
                                                ' Add subsequent captions with left case (default) parameters only
                                                CaptionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                                CaptionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.caption,
                                                                                                 ECloseType.Simple))
                                            End If
                                            CaptionCount += 1
                                        Case TypeConstants.imageSettings

                                            CaptionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))

                                            CaptionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.connection
                                            If FirstStateFound Then
                                                If Not StatesClosed Then
                                                    ' Close off the previous state before starting a connection block
                                                    CaptionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                    CaptionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                                    CaptionFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                                    CaptionFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                                    StatesClosed = True
                                                End If

                                            End If
                                            If FirstConnectionFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                CaptionFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                                CaptionFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                                FirstConnectionFound = True
                                            End If
                                            CaptionFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                            CaptionFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                        Case TypeConstants.state
                                            If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                                ' if this is a subsequent state found after the first then close off the previous state
                                                CaptionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                CaptionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            End If
                                            If FirstStateFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                CaptionFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                                CaptionFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                                FirstStateFound = True
                                            End If
                                            CaptionFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.state,
                                                                                         ECloseType.Complex))
                                            CaptionFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.state,
                                                                                        ECloseType.Complex))
                                            StateCount += 1
                                        Case Else
                                            Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End Select

                                Next
                                If FirstStateFound Then
                                    If Not StatesClosed Then
                                        ' handle the case when no connection block is present and the state blocks need closed
                                        CaptionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        CaptionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                        CaptionFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                        CaptionFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                    End If
                                    StatesClosed = False
                                    FirstStateFound = False
                                    StateCount = 0
                                End If
                                If FirstConnectionFound Then
                                    ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                                    ' We do this at the end of the sub group iteration because we know by observation of the ME software
                                    ' XML object creation that connections always go at the end
                                    CaptionFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                                    CaptionFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                                    FirstConnectionFound = False
                                End If
                                CaptionCount = 0

                                ' Close off this XML object
                                If OType = ECloseType.Complex Then
                                    ' Requires complex object closure
                                    CaptionFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                    CaptionFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                End If
                                TestCount += 1
                                FirstCaptionFound = False
                            Next
                            Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                        End If
                    Next
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                CaptionFileListLeft.Add(itm)
                CaptionFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(CaptionFileListLeft)
            FormatXMLFile(CaptionFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(CaptionFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_caption.xml"))
            WriteOutputFile(CaptionFileListRight, GetPathToLocalFile("Output", FnameVar & "R_caption.xml"))
            WriteOutputFile(CaptionFileCSV, GetPathToLocalFile("Output", FnameVar & "caption.csv"))



            'MsgBox("")

        End If

#End Region

#Region "State List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type = TypeConstants.state Then
                    Type_State_Exists = True
                End If
            Next
        End If


        If Type_State_Exists Then
            ' Generate file content for the main parameter list
            Dim StateFileListLeft As List(Of String) = New List(Of String)
            Dim StateFileListRight As List(Of String) = New List(Of String)
            Dim StateFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            StatesClosed = False
            CaptionCount = 0
            Dim StateMask(10) As Boolean
            Dim StateInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.state)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                StateFileListLeft.Add(itm)
                StateFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            StateFileCSV.Add("Test number,Property,Left,Right")

            ' Set up caption mask
            Select Case True
                Case StateInstCount > 3
                    StateMask(0) = True
                    StateMask(1) = True
                    StateMask(3) = True
                Case StateInstCount > 2
                    StateMask(0) = True
                    StateMask(1) = True
                    StateMask(2) = True
                Case StateInstCount = 1
                    StateMask(0) = True
                Case StateInstCount = 2
                    StateMask(0) = True
                    StateMask(1) = True
                Case StateInstCount = 0
                    StateMask(0) = True
                Case Else
                    Throw New Exception("Whoops, it appears you didnt think of everything")
            End Select

            ' Loop through the test generation process for as many caption test masks are active
            For Smask = 0 To 9
                If StateMask(Smask) Then
                    For Each sublist As subparam In MainObjectList.sList.lSubParamList
                        If sublist.type = TypeConstants.state Then
                            ' This object has a caption sub object type so generate an output file for it
                            For Each sparam As Param In sublist.subParList
                                ' Generate the main objects data with only left cases
                                StateFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                StateFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                'StateFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.caption, TestCount))

                                For Each subp As subparam In MainObjectList.sList.lSubParamList
                                    Select Case subp.type
                                        Case TypeConstants.Data
                                            StateFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                            StateFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.Threshold
                                            StateFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                            StateFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.caption
                                            StateFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         CaptionIndentLevel,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.caption,
                                                                                         ECloseType.Simple))
                                            StateFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        CaptionIndentLevel,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.caption,
                                                                                        ECloseType.Simple))
                                            CaptionCount += 1
                                        Case TypeConstants.imageSettings

                                            StateFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             CaptionIndentLevel,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))

                                            StateFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             CaptionIndentLevel,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.connection
                                            If FirstStateFound Then
                                                If Not StatesClosed Then
                                                    ' Close off the previous state before starting a connection block
                                                    StateFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                    StateFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                                    StateFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                                    StateFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                                    StatesClosed = True
                                                End If

                                            End If
                                            If FirstConnectionFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                StateFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                                StateFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                                FirstConnectionFound = True
                                            End If
                                            StateFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                            StateFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                        Case TypeConstants.state
                                            If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                                ' if this is a subsequent state found after the first then close off the previous state
                                                StateFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                StateFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            End If
                                            If FirstStateFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                StateFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                                StateFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                                FirstStateFound = True
                                            End If
                                            If StateCount = Smask Then ' Select the state instance count to modify based on the mask
                                                ' Check if this parameter is the stateId and skip it, this property value cannot be edited by a user in ME
                                                If Not sparam.sProperty = TypeConstants.stateId Then
                                                    ' Add test case for this state only
                                                    Dim Addstr As String = DetermineAddStrByCase(MainObjectList, StateCount)
                                                    StateFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.state, TestCount, Addstr))
                                                    ' Only substitute params in the first caption object
                                                    StateFileListLeft.Add(CreateXMLObjectStateByDefinition(subp,
                                                                                                           sparam,
                                                                                                           EditCase.Left,
                                                                                                           CaptionIndentLevel - 1,
                                                                                                           TestCount,
                                                                                                           ValuePairList,
                                                                                                           ObjectTestClass.state,
                                                                                                           ECloseType.Complex))

                                                    StateFileListRight.Add(CreateXMLObjectStateByDefinition(subp,
                                                                                                            sparam,
                                                                                                            EditCase.Right,
                                                                                                            CaptionIndentLevel - 1,
                                                                                                            TestCount,
                                                                                                            ValuePairList,
                                                                                                            ObjectTestClass.state,
                                                                                                            ECloseType.Complex))
                                                    FirstStateFound = True
                                                Else
                                                    '' Add subsequent states with left case (default) parameters only
                                                    'StateFileListLeft.Add(CreateXMLObjectByDefinition(subp, EditCase.Left, 1,
                                                    '                                                  TestCount,
                                                    '                                                  ValuePairList,
                                                    '                                                  ObjectTestClass.state,
                                                    '                                                  ECloseType.Complex))
                                                    'StateFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                    '                                                   EditCase.Left,
                                                    '                                                   1,
                                                    '                                                   TestCount,
                                                    '                                                   ValuePairList,
                                                    '                                                   ObjectTestClass.state,
                                                    '                                                   ECloseType.Complex))
                                                    StateFileListLeft.Add(CreateXMLObjectStateByDefinition(subp,
                                                                                                           sparam,
                                                                                                           EditCase.Left,
                                                                                                           CaptionIndentLevel - 1,
                                                                                                           TestCount,
                                                                                                           ValuePairList,
                                                                                                           ObjectTestClass.state,
                                                                                                           ECloseType.Complex))
                                                    StateFileListRight.Add(CreateXMLObjectStateByDefinition(subp,
                                                                                                           sparam,
                                                                                                           EditCase.Left,
                                                                                                           CaptionIndentLevel - 1,
                                                                                                           TestCount,
                                                                                                           ValuePairList,
                                                                                                           ObjectTestClass.state,
                                                                                                           ECloseType.Complex))
                                                End If

                                            Else
                                                ' Add subsequent states with left case (default) parameters only
                                                'StateFileListLeft.Add(CreateXMLObjectByDefinition(subp, EditCase.Left, 1,
                                                '                                                  TestCount,
                                                '                                                  ValuePairList,
                                                '                                                  ObjectTestClass.state,
                                                '                                                  ECloseType.Complex))
                                                'StateFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                '                                                   EditCase.Left,
                                                '                                                   1,
                                                '                                                   TestCount,
                                                '                                                   ValuePairList,
                                                '                                                   ObjectTestClass.state,
                                                '                                                   ECloseType.Complex))
                                                StateFileListLeft.Add(CreateXMLObjectStateByDefinition(subp,
                                                                                                           sparam,
                                                                                                           EditCase.Left,
                                                                                                           CaptionIndentLevel - 1,
                                                                                                           TestCount,
                                                                                                           ValuePairList,
                                                                                                           ObjectTestClass.state,
                                                                                                           ECloseType.Complex))
                                                StateFileListRight.Add(CreateXMLObjectStateByDefinition(subp,
                                                                                                           sparam,
                                                                                                           EditCase.Left,
                                                                                                           CaptionIndentLevel - 1,
                                                                                                           TestCount,
                                                                                                           ValuePairList,
                                                                                                           ObjectTestClass.state,
                                                                                                           ECloseType.Complex))
                                            End If
                                            StateCount += 1
                                        Case Else
                                            Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End Select

                                Next
                                If FirstStateFound Then
                                    If Not StatesClosed Then
                                        ' handle the case when no connection block is present and the state blocks need closed
                                        StateFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        StateFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                        StateFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                        StateFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                    End If
                                    StatesClosed = False
                                    FirstStateFound = False
                                    StateCount = 0
                                End If
                                If FirstConnectionFound Then
                                    ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                                    ' We do this at the end of the sub group iteration because we know by observation of the ME software
                                    ' XML object creation that connections always go at the end
                                    StateFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                                    StateFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                                    FirstConnectionFound = False
                                End If
                                CaptionCount = 0

                                ' Close off this XML object
                                If OType = ECloseType.Complex Then
                                    ' Requires complex object closure
                                    StateFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                    StateFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                End If
                                TestCount += 1
                                FirstCaptionFound = False
                            Next
                            Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                        End If
                    Next
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                StateFileListLeft.Add(itm)
                StateFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(StateFileListLeft)
            FormatXMLFile(StateFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(StateFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_state.xml"))
            WriteOutputFile(StateFileListRight, GetPathToLocalFile("Output", FnameVar & "R_state.xml"))
            WriteOutputFile(StateFileCSV, GetPathToLocalFile("Output", FnameVar & "state.csv"))

            'MsgBox("")

        End If


#End Region

#Region "Connection List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type = TypeConstants.connection Then
                    Type_Connection_Exists = True
                End If
            Next
        End If


        If Type_Connection_Exists Then

            ' Generate file content for the main parameter list
            Dim ConnectionFileListLeft As List(Of String) = New List(Of String)
            Dim ConnectionFileListRight As List(Of String) = New List(Of String)
            Dim ConnectionFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                ConnectionFileListLeft.Add(itm)
                ConnectionFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            ConnectionFileCSV.Add("Test number,Property,Left,Right")

            For Each sublist As subparam In MainObjectList.sList.lSubParamList
                If sublist.type = TypeConstants.connection Then
                    ' This object has an imagesettings sub object type so generate an output file for it
                    For Each sparam As Param In sublist.subParList
                        If Not sparam.sProperty = "name" Then ' Dont process the name property in the connection as it cant be changed by the user
                            ' Generate the main objects data with only left cases
                            ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                            ConnectionFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                            ConnectionFileCSV.Add(CreateTestCaseByConnection(sparam, TestCount, (TestCount - 1)))

                            For Each subp As subparam In MainObjectList.sList.lSubParamList
                                Select Case subp.type
                                    Case TypeConstants.Data
                                        ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.data,
                                                                                         ECloseType.Simple))
                                        ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.data,
                                                                                         ECloseType.Simple))
                                    Case TypeConstants.Threshold
                                        ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.threshold,
                                                                                         ECloseType.Simple))
                                        ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.threshold,
                                                                                         ECloseType.Simple))
                                    Case TypeConstants.state
                                        If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                            ' if this is a subsequent state found after the first then close off the previous state
                                            ConnectionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                            ConnectionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                        End If
                                        If FirstStateFound = False Then
                                            ' Add an additional line here for the connection xml configuration on the first time only
                                            ConnectionFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                            ConnectionFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                            FirstStateFound = True
                                        End If
                                        ConnectionFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.state,
                                                                                     ECloseType.Complex))
                                        ConnectionFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.state,
                                                                                    ECloseType.Complex))
                                        StateCount += 1

                                    Case TypeConstants.caption
                                        ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.caption,
                                                                                         ECloseType.Simple))
                                        ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.caption,
                                                                                         ECloseType.Simple))


                                    Case TypeConstants.imageSettings

                                        ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.image,
                                                                                         ECloseType.Simple))

                                        ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.image,
                                                                                         ECloseType.Simple))
                                    Case TypeConstants.connection
                                        If FirstStateFound Then
                                            If Not StatesClosed Then
                                                ' Close off the previous state before starting a connection block
                                                ConnectionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                ConnectionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                                ConnectionFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                                ConnectionFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                                StatesClosed = True
                                            End If

                                        End If
                                        If FirstConnectionFound = False Then
                                            ' Add an additional line here for the connection xml configuration on the first time only
                                            ConnectionFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                            ConnectionFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                            FirstConnectionFound = True
                                        End If
                                        ' in here we need to select the behaviour of the connection XML generation based on 
                                        ' if this connection matches the current connection selected in the top level for each loop
                                        If sublist.subParList.Item(1).sValue = subp.subParList.Item(1).sValue Then
                                            ' This connection matches the current selected connection at the top level so needs its parameters substituted like
                                            ' The test CSV case
                                            ConnectionFileListLeft.Add(CreateTestXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.caption,
                                                                                         ECloseType.Simple,
                                                                                         ""))
                                            ConnectionFileListRight.Add(CreateTestXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.caption,
                                                                                         ECloseType.Simple,
                                                                                         "s"))
                                        Else
                                            ConnectionFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.caption,
                                                                                         ECloseType.Simple))
                                            ConnectionFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                        End If

                                    Case Else
                                        Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                End Select
                            Next
                            If FirstStateFound Then
                                If Not StatesClosed Then
                                    ' handle the case when no connection block is present and the state blocks need closed
                                    ConnectionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                    ConnectionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                    ConnectionFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                    ConnectionFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                End If
                                StatesClosed = False
                                FirstStateFound = False
                                StateCount = 0
                            End If
                            If FirstConnectionFound Then
                                ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                                ' We do this at the end of the sub group iteration because we know by observation of the ME software
                                ' XML object creation that connections always go at the end
                                ConnectionFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                                ConnectionFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                                FirstConnectionFound = False
                            End If

                            ' Close off this XML object
                            If OType = ECloseType.Complex Then
                                ' Requires complex object closure
                                ConnectionFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                ConnectionFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                            End If
                            TestCount += 1
                        End If
                    Next
                End If
            Next

            ' Handle connection list special cases
            ' Count number of connections present in sublist
            ' Also check if any of the connections has an optional expression
            Dim ConnectionCount As Integer = 0
            Dim HasOptionalExpression As Boolean = False
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type = TypeConstants.connection Then
                    ConnectionCount += 1
                End If
                For Each subitm As Param In itm.subParList
                    If subitm.sProperty = TypeConstants.optionalExpression Then
                        HasOptionalExpression = True
                    End If
                Next
            Next

            Dim SpecialCaseMessage As String = ""
            If ConnectionCount > 0 Then
                SpecialCaseMessage = InputBox("Enter the name of the object type as it will appear" & vbCrLf &
                                              "in the test CSV for connection count mismatch", "Special case message string req", "")
            End If

            Select Case HasOptionalExpression
                Case True
                    Throw New Exception("Oops, not handled yet, get yer finger oot!")
                Case False
                    Select Case ConnectionCount
                        Case 0
                        ' Do nothing in here, this object has no connections
                        Case 1

                            For SpecialCases = 1 To 2
                                Call GenerateXMLObjectWithoutConnections(MainObjectList,
                                                                     ConnectionFileListLeft,
                                                                     ConnectionFileListRight,
                                                                     ConnectionFileCSV,
                                                                     TestCount,
                                                                     ValuePairList,
                                                                     FirstConnectionFound,
                                                                     OType)
                                Select Case SpecialCases
                                    Case 1 ' Left no connections, right defined
                                        Call AddConnectionsNoParams(MainObjectList, ConnectionFileListRight, ValuePairList)
                                        ' Dont add the left here as its missing in this case

                                        ' Add test case to the CSV file
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial(SpecialCaseMessage & " - Connection Count Mismatch",
                                                                                                "nothing",
                                                                                                "defined",
                                                                                                TestCount))
                                    Case 2 ' right no connections, left defined
                                        Call AddConnectionsNoParams(MainObjectList, ConnectionFileListLeft, ValuePairList)
                                        ' Dont add the right here as its missing in this case
                                        ' Add test case to the CSV file
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial(SpecialCaseMessage & " - Connection Count Mismatch",
                                                                                                "defined",
                                                                                                "nothing",
                                                                                                TestCount))
                                End Select

                                ' Close off this XML object
                                If OType = ECloseType.Complex Then
                                    ' Requires complex object closure
                                    ConnectionFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                    ConnectionFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                End If

                                TestCount += 1

                            Next


                        Case Else
                            ' More than 1 connection so generate the usual test cases

                            For SpecialCases = 1 To 6
                                Call GenerateXMLObjectWithoutConnections(MainObjectList,
                                                                     ConnectionFileListLeft,
                                                                     ConnectionFileListRight,
                                                                     ConnectionFileCSV,
                                                                     TestCount,
                                                                     ValuePairList,
                                                                     FirstConnectionFound,
                                                                     OType)
                                Select Case SpecialCases
                                    Case 1 ' left no connections - right defined
                                        Call AddConnectionsNoParams(MainObjectList, ConnectionFileListRight, ValuePairList)
                                        ' Dont add the left here as its missing in this case

                                        ' Add test case to the CSV file
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial(SpecialCaseMessage & " - Connection Count Mismatch",
                                                                                                "nothing",
                                                                                                "defined",
                                                                                                TestCount))

                                    Case 2 ' right no connections - left defined
                                        Call AddConnectionsNoParams(MainObjectList, ConnectionFileListLeft, ValuePairList)
                                        ' Dont add the right here as its missing in this case
                                        ' Add test case to the CSV file
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial(SpecialCaseMessage & " - Connection Count Mismatch",
                                                                                                "defined",
                                                                                                "nothing",
                                                                                                TestCount))
                                    Case 3 ' left - right connection count mismatch, 1 conn missing from left
                                        ' Add entries into both lists but filter 1 parameter each
                                        Dim LeftParStr As String = GetTestCaseValueForConnectionBySubparam(MainObjectList, 2)
                                        Dim RightParStr As String = GetTestCaseValueForConnectionBySubparam(MainObjectList, 1)
                                        Call AddConnectionsFilterParam(MainObjectList, ConnectionFileListLeft, ValuePairList, 1)
                                        Call AddConnectionsFilterParam(MainObjectList, ConnectionFileListRight, ValuePairList, 2)
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial("Connection 0 - expression",
                                                                                                LeftParStr,
                                                                                                RightParStr,
                                                                                                TestCount))
                                    Case 4 ' right - left connection count mismatch, 1 conn missing from right
                                        ' Add entries into both lists but filter 1 parameter each
                                        Dim LeftParStr As String = GetTestCaseValueForConnectionBySubparam(MainObjectList, 1)
                                        Dim RightParStr As String = GetTestCaseValueForConnectionBySubparam(MainObjectList, 2)
                                        Call AddConnectionsFilterParam(MainObjectList, ConnectionFileListLeft, ValuePairList, 2)
                                        Call AddConnectionsFilterParam(MainObjectList, ConnectionFileListRight, ValuePairList, 1)
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial("Connection 0 - expression",
                                                                                                LeftParStr,
                                                                                                RightParStr,
                                                                                                TestCount))
                                    Case 5 ' left <> right, each side has 1 connection missing but different for each side
                                        ' Add all to right side, skip 1 on the left
                                        Call AddConnectionsNoParams(MainObjectList, ConnectionFileListRight, ValuePairList)
                                        Call AddConnectionsFilterParam(MainObjectList, ConnectionFileListLeft, ValuePairList, 1)
                                        ' Add test case to the CSV file
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial(SpecialCaseMessage & " - Connection Count Mismatch",
                                                                                                (ConnectionCount - 1).ToString,
                                                                                                ConnectionCount.ToString,
                                                                                                TestCount))
                                    Case 6 ' right <> left, each side has 1 connection missing but different for each side
                                        ' Add all to left side, skip 1 on the right
                                        Call AddConnectionsNoParams(MainObjectList, ConnectionFileListLeft, ValuePairList)
                                        Call AddConnectionsFilterParam(MainObjectList, ConnectionFileListRight, ValuePairList, 1)
                                        ' Add test case to the CSV file
                                        ConnectionFileCSV.Add(CreateTestCaseByConnectionSpecial(SpecialCaseMessage & " - Connection Count Mismatch",
                                                                                                ConnectionCount.ToString,
                                                                                                (ConnectionCount - 1).ToString,
                                                                                                TestCount))
                                End Select

                                ' Close off this XML object
                                If OType = ECloseType.Complex Then
                                    ' Requires complex object closure
                                    ConnectionFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                    ConnectionFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                End If

                                TestCount += 1

                            Next

                    End Select
            End Select

            ' Close off the files with the footer
            For Each itm As String In FooterList
                ConnectionFileListLeft.Add(itm)
                ConnectionFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(ConnectionFileListLeft)
            FormatXMLFile(ConnectionFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(ConnectionFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_connections.xml"))
            WriteOutputFile(ConnectionFileListRight, GetPathToLocalFile("Output", FnameVar & "R_connections.xml"))
            WriteOutputFile(ConnectionFileCSV, GetPathToLocalFile("Output", FnameVar & "connections.csv"))

            'MsgBox("")

        End If

#End Region

#Region "Threshold List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type = TypeConstants.Threshold Then
                    Type_Threshold_Exists = True
                End If
            Next
        End If


        If Type_Threshold_Exists Then
            ' Generate file content for the main parameter list
            Dim ThresholdFileListLeft As List(Of String) = New List(Of String)
            Dim ThresholdFileListRight As List(Of String) = New List(Of String)
            Dim ThresholdFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False
            FirstCaptionFound = False ' Added to ensure only the first caption type gets processed when dealing with mutlistate objects
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            CaptionCount = 0
            ThresholdCount = 0
            Dim ThresholdMask(10) As Boolean
            Dim ThresholdInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.Threshold)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                ThresholdFileListLeft.Add(itm)
                ThresholdFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            ThresholdFileCSV.Add("Test number,Property,Left,Right")

            ' Set up caption mask
            Select Case ThresholdInstCount
                Case 1
                    ThresholdMask(0) = True
                    ThresholdMask(1) = False
                Case 2
                    ThresholdMask(0) = True
                    ThresholdMask(1) = True
                Case Else
                    Throw New Exception("Whoops, it appears you didnt think of everything")
            End Select

            ' Loop through the test generation process for as many caption test masks are active
            For Tmask = 0 To 9
                If ThresholdMask(Tmask) Then
                    For Each sublist As subparam In MainObjectList.sList.lSubParamList
                        If sublist.type = TypeConstants.Threshold Then
                            ' This object has a caption sub object type so generate an output file for it
                            For Each sparam As Param In sublist.subParList
                                ' Generate the main objects data with only left cases
                                ThresholdFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                ThresholdFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                'ThresholdFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.caption, TestCount))

                                For Each subp As subparam In MainObjectList.sList.lSubParamList
                                    Select Case subp.type
                                        Case TypeConstants.Data
                                            ThresholdFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                            ThresholdFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.Threshold
                                            If ThresholdCount = Tmask Then
                                                If Not sparam.sProperty = "thresholdIndex" Then
                                                    Dim addstr As String = "Threshold " & ThresholdCount & " - "
                                                    ThresholdFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.threshold, TestCount, addstr))
                                                    ThresholdFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Left,
                                                                                                  1,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.threshold,
                                                                                                  ECloseType.Simple))
                                                    ThresholdFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Right,
                                                                                                  1,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.threshold,
                                                                                                  ECloseType.Simple))
                                                Else
                                                    ThresholdFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                                    ThresholdFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                     EditCase.Left,
                                                                                                     1,
                                                                                                     TestCount,
                                                                                                     ValuePairList,
                                                                                                     ObjectTestClass.threshold,
                                                                                                     ECloseType.Simple))
                                                End If

                                            Else
                                                ThresholdFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                                ThresholdFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.threshold,
                                                                                                 ECloseType.Simple))
                                            End If
                                            ThresholdCount += 1

                                        Case TypeConstants.caption
                                            ThresholdFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                            ThresholdFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                            CaptionCount += 1
                                        Case TypeConstants.imageSettings

                                            ThresholdFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))

                                            ThresholdFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.connection
                                            If FirstStateFound Then
                                                If Not StatesClosed Then
                                                    ' Close off the previous state before starting a connection block
                                                    ThresholdFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                    ThresholdFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                                    ThresholdFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                                    ThresholdFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                                    StatesClosed = True
                                                End If

                                            End If
                                            If FirstConnectionFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                ThresholdFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                                ThresholdFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                                FirstConnectionFound = True
                                            End If
                                            ThresholdFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                            ThresholdFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                        Case TypeConstants.state
                                            If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                                ' if this is a subsequent state found after the first then close off the previous state
                                                ThresholdFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                ThresholdFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            End If
                                            If FirstStateFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                ThresholdFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                                ThresholdFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                                FirstStateFound = True
                                            End If
                                            ThresholdFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.state,
                                                                                         ECloseType.Complex))
                                            ThresholdFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.state,
                                                                                        ECloseType.Complex))
                                            StateCount += 1
                                        Case Else
                                            Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End Select

                                Next
                                If FirstStateFound Then
                                    If Not StatesClosed Then
                                        ' handle the case when no connection block is present and the state blocks need closed
                                        ThresholdFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        ThresholdFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                        ThresholdFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                        ThresholdFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                    End If
                                    StatesClosed = False
                                    FirstStateFound = False
                                    StateCount = 0
                                End If
                                If FirstConnectionFound Then
                                    ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                                    ' We do this at the end of the sub group iteration because we know by observation of the ME software
                                    ' XML object creation that connections always go at the end
                                    ThresholdFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                                    ThresholdFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                                    FirstConnectionFound = False
                                End If
                                CaptionCount = 0
                                ThresholdCount = 0

                                ' Close off this XML object
                                If OType = ECloseType.Complex Then
                                    ' Requires complex object closure
                                    ThresholdFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                    ThresholdFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                End If
                                TestCount += 1
                                FirstCaptionFound = False
                            Next
                            Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                        End If
                    Next
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                ThresholdFileListLeft.Add(itm)
                ThresholdFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(ThresholdFileListLeft)
            FormatXMLFile(ThresholdFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(ThresholdFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_threshold.xml"))
            WriteOutputFile(ThresholdFileListRight, GetPathToLocalFile("Output", FnameVar & "R_threshold.xml"))
            WriteOutputFile(ThresholdFileCSV, GetPathToLocalFile("Output", FnameVar & "threshold.csv"))



            MsgBox("")

        End If

#End Region

#Region "Animations List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type Like TypeConstants.Animate & "*" Then
                    Type_Animation_Exists = True
                End If
            Next
        End If


        If Type_Animation_Exists Then
            ' Generate file content for the main parameter list
            Dim AnimationFileListLeft As List(Of String) = New List(Of String)
            Dim AnimationFileListRight As List(Of String) = New List(Of String)
            Dim AnimationFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False
            FirstCaptionFound = False ' Added to ensure only the first caption type gets processed when dealing with mutlistate objects
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            CaptionCount = 0
            ThresholdCount = 0
            FirstAnimationFound = 0
            AnimationsClosed = False
            Dim ThresholdMask(10) As Boolean
            Dim ThresholdInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.Threshold)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                AnimationFileListLeft.Add(itm)
                AnimationFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            AnimationFileCSV.Add("Test number,Property,Left,Right")

            ' Loop through the test generation process for as many caption test masks are active
            For Each sublist As subparam In MainObjectList.sList.lSubParamList
                If sublist.type Like TypeConstants.Animate & "*" Then
                    ' This object has an animation sub object type so generate an output file for it
                    For Each sparam As Param In sublist.subParList
                        ' Generate the main objects data with only left cases
                        AnimationFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                        AnimationFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))

                        For Each subp As subparam In MainObjectList.sList.lSubParamList
                            Select Case subp.type
                                Case TypeConstants.Data
                                    AnimationFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                    AnimationFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                Case TypeConstants.Threshold
                                    AnimationFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                    AnimationFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.threshold,
                                                                                                 ECloseType.Simple))
                                    ThresholdCount += 1

                                Case TypeConstants.caption
                                    AnimationFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                    AnimationFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                    CaptionCount += 1
                                Case TypeConstants.imageSettings

                                    AnimationFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))

                                    AnimationFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                Case TypeConstants.connection
                                    If FirstAnimationFound Then
                                        If Not AnimationsClosed Then
                                            If FirstColorFound Then
                                                'If Not ColorsClosed Then
                                                '    ' close off the animatecolor block
                                                '    MainFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                                '    MainFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                                '    ColorsClosed = True
                                                'End If
                                                If Not TagToClose = "" Then
                                                    MainFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    MainFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    TagToClose = ""
                                                End If
                                            End If
                                            ' Add lines here to close off the animation objects
                                            MainFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                            MainFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                            AnimationsClosed = True
                                        End If
                                    End If
                                    If FirstStateFound Then
                                        If Not StatesClosed Then
                                            ' Close off the previous state before starting a connection block
                                            AnimationFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                            AnimationFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            AnimationFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                            AnimationFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                            StatesClosed = True
                                        End If

                                    End If
                                    If FirstConnectionFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        AnimationFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                        AnimationFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                        FirstConnectionFound = True
                                    End If
                                    AnimationFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                    AnimationFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                Case TypeConstants.state
                                    If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                        ' if this is a subsequent state found after the first then close off the previous state
                                        AnimationFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        AnimationFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                    End If
                                    If FirstStateFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        AnimationFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                        AnimationFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                        FirstStateFound = True
                                    End If
                                    AnimationFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.state,
                                                                                         ECloseType.Complex))
                                    AnimationFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.state,
                                                                                        ECloseType.Complex))
                                    StateCount += 1
                                Case TypeConstants.Color
                                    If FirstColorFound = False Then
                                        FirstColorFound = True
                                    End If
                                    AnimationFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.color,
                                                                                     ECloseType.Simple))
                                    AnimationFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.color,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.readFromTagExpressionRange
                                    AnimationFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.readfromtagexpressionrange,
                                                                                     ECloseType.Simple))
                                    AnimationFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.readfromtagexpressionrange,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.constantExpressionRange
                                    AnimationFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.constantexpressionrange,
                                                                                     ECloseType.Simple))
                                    AnimationFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.constantexpressionrange,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.defaultExpressionRange
                                    AnimationFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.defaultexpressionrange,
                                                                                     ECloseType.Simple))
                                    AnimationFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.defaultexpressionrange,
                                                                                    ECloseType.Simple))
                                Case Else
                                    If subp.type Like TypeConstants.Animate & "*" Then
                                        If FirstAnimationFound = False Then
                                            ' Add an additional line here for the connection xml configuration on the first time only
                                            AnimationFileListLeft.Add(AddWhiteSpace(1, "<animations>"))
                                            AnimationFileListRight.Add(AddWhiteSpace(1, "<animations>"))
                                            FirstAnimationFound = True
                                        End If
                                        Dim Addstr As String = GetAddStrByAnimationType(sublist.type)
                                        Dim SelectEcloseType As ECloseType = GetAnimationEcloseType(subp.type)
                                        If SelectEcloseType = ECloseType.Complex Then
                                            ' Store the name of the animation tag so we can close it later
                                            ' Also check if the tag to store has changed so we can close previous tags
                                            If Not TagToClose = "" Then
                                                ' case when closing tag already exists
                                                If TagToClose = subp.type Then
                                                    ' Do nothing because this is the same type as before, it will be closed at the end
                                                    ' Of the main loop
                                                Else
                                                    ' It is a different type, close out the old one and start a new tag
                                                    AnimationFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    AnimationFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    TagToClose = subp.type
                                                    AnimationFileListLeft.Add(AddWhiteSpace(1, "<" & TagToClose & ">"))
                                                    AnimationFileListRight.Add(AddWhiteSpace(1, "<" & TagToClose & ">"))
                                                End If
                                            Else
                                                ' tagtoclose not set yet so update it with current type
                                                TagToClose = subp.type
                                            End If
                                        Else
                                            ' Upon encountering a simple type check if a previous complex type needs closed first
                                            If Not TagToClose = "" Then
                                                ' a previous tag needs closed first
                                                AnimationFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                AnimationFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                TagToClose = "" ' this lets us know on the next loop that nothing requires closing
                                            Else
                                                ' Do nothing
                                                ' no tags opened to be closed and this is a simple type so just deal with it normally
                                            End If
                                        End If
                                        AnimationFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.animate, TestCount, Addstr))
                                        AnimationFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                              sparam,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.animate,
                                                                                                 SelectEcloseType))

                                        AnimationFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                              sparam,
                                                                                                 EditCase.Right,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.animate,
                                                                                                 SelectEcloseType))
                                    Else
                                        Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End If
                            End Select

                        Next
                        If FirstAnimationFound Then
                            If Not AnimationsClosed Then
                                'If FirstColorFound Then
                                '    If Not ColorsClosed Then
                                '        ' close off the animatecolor block
                                '        AnimationFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                '        AnimationFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                '        ColorsClosed = True
                                '    End If
                                'End If
                                If Not TagToClose = "" Then
                                    AnimationFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                    AnimationFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                    TagToClose = ""
                                End If
                                ' Add lines here to close off the animation objects
                                AnimationFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                AnimationFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                AnimationsClosed = True
                            End If
                        End If
                        FirstAnimationFound = False
                        AnimationsClosed = False
                        ColorsClosed = False
                        FirstColorFound = False
                        If FirstStateFound Then
                            If Not StatesClosed Then
                                ' handle the case when no connection block is present and the state blocks need closed
                                AnimationFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                AnimationFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                AnimationFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                AnimationFileListRight.Add(AddWhiteSpace(1, "</states>"))
                            End If
                            StatesClosed = False
                            FirstStateFound = False
                            StateCount = 0
                        End If
                        If FirstConnectionFound Then
                            ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                            ' We do this at the end of the sub group iteration because we know by observation of the ME software
                            ' XML object creation that connections always go at the end
                            AnimationFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                            AnimationFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                            FirstConnectionFound = False
                        End If
                        CaptionCount = 0
                        ThresholdCount = 0

                        ' Close off this XML object
                        If OType = ECloseType.Complex Then
                            ' Requires complex object closure
                            AnimationFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                            AnimationFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                        End If
                        TestCount += 1
                        FirstCaptionFound = False
                    Next
                    ' This exit for is removed as we want to generate content for all animation type objects found in the test
                    'Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                AnimationFileListLeft.Add(itm)
                AnimationFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(AnimationFileListLeft)
            FormatXMLFile(AnimationFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(AnimationFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_animate.xml"))
            WriteOutputFile(AnimationFileListRight, GetPathToLocalFile("Output", FnameVar & "R_animate.xml"))
            WriteOutputFile(AnimationFileCSV, GetPathToLocalFile("Output", FnameVar & "animate.csv"))



            MsgBox("")

        End If

#End Region

#Region "Animations Expression Range  List Generation"
        Dim ExpressionRangeType As String = ""
        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                Select Case True
                    ' Disable default type as this contains no testable parameters
                    'Case itm.type = TypeConstants.defaultExpressionRange
                    '    Type_ExpressionRange_Exists = True
                    Case itm.type = TypeConstants.readFromTagExpressionRange
                        Type_ExpressionRange_Exists = True
                        ExpressionRangeType = "readfromtag"
                    Case itm.type = TypeConstants.constantExpressionRange
                        Type_ExpressionRange_Exists = True
                        ExpressionRangeType = "constant"
                End Select
            Next
        End If

        Dim SpecialCaseAnimationMessage As String = ""
        SpecialCaseAnimationMessage = InputBox("Enter the name of the object animation type message as it will appear" & vbCrLf &
                                              "in the test CSV for expression range animations", "Special case message string req", "")


        If Type_ExpressionRange_Exists Then
            ' Generate file content for the main parameter list
            Dim ExpressionRangeFileListLeft As List(Of String) = New List(Of String)
            Dim ExpressionRangeFileListRight As List(Of String) = New List(Of String)
            Dim ExpressionRangeFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False
            FirstCaptionFound = False ' Added to ensure only the first caption type gets processed when dealing with mutlistate objects
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            CaptionCount = 0
            ThresholdCount = 0
            FirstAnimationFound = 0
            AnimationsClosed = False
            Dim ThresholdMask(10) As Boolean
            Dim ThresholdInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.Threshold)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                ExpressionRangeFileListLeft.Add(itm)
                ExpressionRangeFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            ExpressionRangeFileCSV.Add("Test number,Property,Left,Right")

            ' Loop through the test generation process for as many caption test masks are active
            For Each sublist As subparam In MainObjectList.sList.lSubParamList
                If sublist.type Like "*ExpressionRange" Then
                    ' This object has an animation sub object type so generate an output file for it
                    For Each sparam As Param In sublist.subParList
                        ' Generate the main objects data with only left cases
                        ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                        ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))

                        For Each subp As subparam In MainObjectList.sList.lSubParamList
                            Select Case subp.type
                                Case TypeConstants.Data
                                    ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                    ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                Case TypeConstants.Threshold
                                    ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                    ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.threshold,
                                                                                                 ECloseType.Simple))
                                    ThresholdCount += 1

                                Case TypeConstants.caption
                                    ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                    ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                    CaptionCount += 1
                                Case TypeConstants.imageSettings

                                    ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))

                                    ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                Case TypeConstants.connection
                                    If FirstAnimationFound Then
                                        If Not AnimationsClosed Then
                                            If FirstColorFound Then
                                                'If Not ColorsClosed Then
                                                '    ' close off the animatecolor block
                                                '    MainFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                                '    MainFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                                '    ColorsClosed = True
                                                'End If
                                                If Not TagToClose = "" Then
                                                    MainFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    MainFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    TagToClose = ""
                                                End If
                                            End If
                                            ' Add lines here to close off the animation objects
                                            MainFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                            MainFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                            AnimationsClosed = True
                                        End If
                                    End If
                                    If FirstStateFound Then
                                        If Not StatesClosed Then
                                            ' Close off the previous state before starting a connection block
                                            ExpressionRangeFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                            ExpressionRangeFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                            ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                            StatesClosed = True
                                        End If

                                    End If
                                    If FirstConnectionFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                        ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                        FirstConnectionFound = True
                                    End If
                                    ExpressionRangeFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                    ExpressionRangeFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                Case TypeConstants.state
                                    If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                        ' if this is a subsequent state found after the first then close off the previous state
                                        ExpressionRangeFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        ExpressionRangeFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                    End If
                                    If FirstStateFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                        ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                        FirstStateFound = True
                                    End If
                                    ExpressionRangeFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.state,
                                                                                         ECloseType.Complex))
                                    ExpressionRangeFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.state,
                                                                                        ECloseType.Complex))
                                    StateCount += 1
                                Case TypeConstants.Color
                                    If FirstColorFound = False Then
                                        FirstColorFound = True
                                    End If
                                    ExpressionRangeFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     2,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.color,
                                                                                     ECloseType.Simple))
                                    ExpressionRangeFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                    EditCase.Left,
                                                                                    2,
                                                                                    TestCount,
                                                                                    ValuePairList,
                                                                                    ObjectTestClass.color,
                                                                                    ECloseType.Simple))
                                Case TypeConstants.readFromTagExpressionRange
                                    Dim addstr As String = GetAddStrByAnimationType(sublist.type)
                                    addstr = SpecialCaseAnimationMessage + addstr
                                    ExpressionRangeFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.readfromtagexpressionrange, TestCount, addstr))
                                    ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Left,
                                                                                                  1,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.readfromtagexpressionrange,
                                                                                                  ECloseType.Simple))
                                    ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Right,
                                                                                                  1,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.readfromtagexpressionrange,
                                                                                                  ECloseType.Simple))
                                Case TypeConstants.constantExpressionRange
                                    Dim addstr As String = GetAddStrByAnimationType(sublist.type)
                                    addstr = SpecialCaseAnimationMessage + addstr
                                    ExpressionRangeFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.constantexpressionrange, TestCount, addstr))
                                    ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Left,
                                                                                                  1,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.constantexpressionrange,
                                                                                                  ECloseType.Simple))
                                    ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                  sparam,
                                                                                                  EditCase.Right,
                                                                                                  1,
                                                                                                  TestCount,
                                                                                                  ValuePairList,
                                                                                                  ObjectTestClass.constantexpressionrange,
                                                                                                  ECloseType.Simple))
                                Case Else
                                    If subp.type Like TypeConstants.Animate & "*" Then
                                        If FirstAnimationFound = False Then
                                            ' Add an additional line here for the connection xml configuration on the first time only
                                            ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "<animations>"))
                                            ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "<animations>"))
                                            FirstAnimationFound = True
                                        End If
                                        Dim Addstr As String = GetAddStrByAnimationType(sublist.type)
                                        Dim SelectEcloseType As ECloseType = GetAnimationEcloseType(subp.type)
                                        If SelectEcloseType = ECloseType.Complex Then
                                            ' Store the name of the animation tag so we can close it later
                                            ' Also check if the tag to store has changed so we can close previous tags
                                            If Not TagToClose = "" Then
                                                ' case when closing tag already exists
                                                If TagToClose = subp.type Then
                                                    ' Do nothing because this is the same type as before, it will be closed at the end
                                                    ' Of the main loop
                                                Else
                                                    ' It is a different type, close out the old one and start a new tag
                                                    ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                    TagToClose = subp.type
                                                    ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "<" & TagToClose & ">"))
                                                    ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "<" & TagToClose & ">"))
                                                End If
                                            Else
                                                ' tagtoclose not set yet so update it with current type
                                                TagToClose = subp.type
                                            End If
                                        Else
                                            ' Upon encountering a simple type check if a previous complex type needs closed first
                                            If Not TagToClose = "" Then
                                                ' a previous tag needs closed first
                                                ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                                TagToClose = "" ' this lets us know on the next loop that nothing requires closing
                                            Else
                                                ' Do nothing
                                                ' no tags opened to be closed and this is a simple type so just deal with it normally
                                            End If
                                        End If
                                        'ExpressionRangeFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.animate, TestCount, Addstr))
                                        ExpressionRangeFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.animate,
                                                                                                 SelectEcloseType))

                                        ExpressionRangeFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.animate,
                                                                                                 SelectEcloseType))
                                    Else
                                        Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End If
                            End Select

                        Next
                        If FirstAnimationFound Then
                            If Not AnimationsClosed Then
                                'If FirstColorFound Then
                                '    If Not ColorsClosed Then
                                '        ' close off the animatecolor block
                                '        ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                '        ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                '        ColorsClosed = True
                                '    End If
                                'End If
                                If Not TagToClose = "" Then
                                    ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                    ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</" & TagToClose & ">"))
                                    TagToClose = ""
                                End If
                                ' Add lines here to close off the animation objects
                                ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                AnimationsClosed = True
                            End If
                        End If
                        FirstAnimationFound = False
                        AnimationsClosed = False
                        ColorsClosed = False
                        FirstColorFound = False
                        If FirstStateFound Then
                            If Not StatesClosed Then
                                ' handle the case when no connection block is present and the state blocks need closed
                                ExpressionRangeFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                ExpressionRangeFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</states>"))
                            End If
                            StatesClosed = False
                            FirstStateFound = False
                            StateCount = 0
                        End If
                        If FirstConnectionFound Then
                            ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                            ' We do this at the end of the sub group iteration because we know by observation of the ME software
                            ' XML object creation that connections always go at the end
                            ExpressionRangeFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                            ExpressionRangeFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                            FirstConnectionFound = False
                        End If
                        CaptionCount = 0
                        ThresholdCount = 0

                        ' Close off this XML object
                        If OType = ECloseType.Complex Then
                            ' Requires complex object closure
                            ExpressionRangeFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                            ExpressionRangeFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                        End If
                        TestCount += 1
                        FirstCaptionFound = False
                    Next
                    ' This exit for is removed as we want to generate content for all animation type objects found in the test
                    'Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                ExpressionRangeFileListLeft.Add(itm)
                ExpressionRangeFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(ExpressionRangeFileListLeft)
            FormatXMLFile(ExpressionRangeFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            ' Use variable names here to allow for output of different file names for other test types
            WriteOutputFile(ExpressionRangeFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_animate" & ExpressionRangeType & ".xml"))
            WriteOutputFile(ExpressionRangeFileListRight, GetPathToLocalFile("Output", FnameVar & "R_animate" & ExpressionRangeType & ".xml"))
            WriteOutputFile(ExpressionRangeFileCSV, GetPathToLocalFile("Output", FnameVar & "animate" & ExpressionRangeType & ".csv"))



            MsgBox("")

        End If

#End Region

#Region "Color List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type Like TypeConstants.Color & "*" Then
                    Type_Color_Exists = True
                End If
            Next
        End If


        If Type_Color_Exists Then
            ' Generate file content for the main parameter list
            Dim AnimationColorFileListLeft As List(Of String) = New List(Of String)
            Dim AnimationColorFileListRight As List(Of String) = New List(Of String)
            Dim AnimationColorFileCSV As List(Of String) = New List(Of String)
            Dim ColorCount As Integer = 0
            TestCount = 1
            FirstConnectionFound = False
            FirstCaptionFound = False ' Added to ensure only the first caption type gets processed when dealing with mutlistate objects
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            CaptionCount = 0
            ThresholdCount = 0
            FirstAnimationFound = 0
            AnimationsClosed = False
            Dim ThresholdMask(10) As Boolean
            Dim ThresholdInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.Threshold)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                AnimationColorFileListLeft.Add(itm)
                AnimationColorFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            AnimationColorFileCSV.Add("Test number,Property,Left,Right")

            ' Loop through the test generation process for as many caption test masks are active
            For Each sublist As subparam In MainObjectList.sList.lSubParamList
                If sublist.type Like TypeConstants.Color Then
                    ' This object has an animation sub object type so generate an output file for it
                    For Each sparam As Param In sublist.subParList
                        ' Generate the main objects data with only left cases
                        AnimationColorFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                        AnimationColorFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))

                        For Each subp As subparam In MainObjectList.sList.lSubParamList
                            Select Case subp.type
                                Case TypeConstants.Data
                                    AnimationColorFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                    AnimationColorFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                Case TypeConstants.Threshold
                                    AnimationColorFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))
                                    AnimationColorFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.threshold,
                                                                                                 ECloseType.Simple))
                                    ThresholdCount += 1

                                Case TypeConstants.caption
                                    AnimationColorFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                    AnimationColorFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                    CaptionCount += 1
                                Case TypeConstants.imageSettings

                                    AnimationColorFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))

                                    AnimationColorFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                Case TypeConstants.connection
                                    If FirstAnimationFound Then
                                        If Not AnimationsClosed Then
                                            If FirstColorFound Then
                                                If Not ColorsClosed Then
                                                    ' close off the animatecolor block
                                                    MainFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                                    MainFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                                    ColorsClosed = True
                                                End If
                                            End If
                                            ' Add lines here to close off the animation objects
                                            MainFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                            MainFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                            AnimationsClosed = True
                                        End If
                                    End If
                                    If FirstStateFound Then
                                        If Not StatesClosed Then
                                            ' Close off the previous state before starting a connection block
                                            AnimationColorFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                            AnimationColorFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            AnimationColorFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                            AnimationColorFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                            StatesClosed = True
                                        End If

                                    End If
                                    If FirstConnectionFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        AnimationColorFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                        AnimationColorFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                        FirstConnectionFound = True
                                    End If
                                    AnimationColorFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                    AnimationColorFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                Case TypeConstants.state
                                    If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                        ' if this is a subsequent state found after the first then close off the previous state
                                        AnimationColorFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        AnimationColorFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                    End If
                                    If FirstStateFound = False Then
                                        ' Add an additional line here for the connection xml configuration on the first time only
                                        AnimationColorFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                        AnimationColorFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                        FirstStateFound = True
                                    End If
                                    AnimationColorFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.state,
                                                                                         ECloseType.Complex))
                                    AnimationColorFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.state,
                                                                                        ECloseType.Complex))
                                    StateCount += 1
                                Case TypeConstants.Color
                                    If FirstColorFound = False Then
                                        FirstColorFound = True
                                    End If
                                    ' Add test case for this image only
                                    Dim Addstr As String = GetAddStrByAnimationType(sublist.type)
                                    Addstr &= ColorCount.ToString & " - "
                                    AnimationColorFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.color, TestCount, Addstr, subp.type))
                                    ' Only substitute params in the first image object
                                    AnimationColorFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                      sparam,
                                                                                      EditCase.Left,
                                                                                      CaptionIndentLevel,
                                                                                      TestCount,
                                                                                      ValuePairList,
                                                                                      ObjectTestClass.color,
                                                                                      ECloseType.Simple))

                                    AnimationColorFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                          sparam,
                                                                                          EditCase.Right,
                                                                                          CaptionIndentLevel,
                                                                                          TestCount,
                                                                                          ValuePairList,
                                                                                          ObjectTestClass.color,
                                                                                          ECloseType.Simple))
                                    ColorCount += 1
                                Case Else
                                    If subp.type Like TypeConstants.Animate & "*" Then
                                        If FirstAnimationFound = False Then
                                            ' Add an additional line here for the connection xml configuration on the first time only
                                            AnimationColorFileListLeft.Add(AddWhiteSpace(1, "<animations>"))
                                            AnimationColorFileListRight.Add(AddWhiteSpace(1, "<animations>"))
                                            FirstAnimationFound = True
                                        End If
                                        Dim SelectEcloseType As ECloseType = GetAnimationEcloseType(subp.type)
                                        AnimationColorFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.animate,
                                                                                         SelectEcloseType))
                                        AnimationColorFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         1,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.animate,
                                                                                         SelectEcloseType))
                                    Else
                                        Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End If
                            End Select
                        Next
                        ColorCount = 0
                        If FirstAnimationFound Then
                            If Not AnimationsClosed Then
                                If FirstColorFound Then
                                    If Not ColorsClosed Then
                                        ' close off the animatecolor block
                                        AnimationColorFileListLeft.Add(AddWhiteSpace(1, "</animateColor>"))
                                        AnimationColorFileListRight.Add(AddWhiteSpace(1, "</animateColor>"))
                                        ColorsClosed = True
                                    End If
                                End If
                                ' Add lines here to close off the animation objects
                                AnimationColorFileListLeft.Add(AddWhiteSpace(1, "</animations>"))
                                AnimationColorFileListRight.Add(AddWhiteSpace(1, "</animations>"))
                                AnimationsClosed = True
                            End If
                        End If
                        FirstAnimationFound = False
                        AnimationsClosed = False
                        ColorsClosed = False
                        FirstColorFound = False
                        If FirstStateFound Then
                            If Not StatesClosed Then
                                ' handle the case when no connection block is present and the state blocks need closed
                                AnimationColorFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                AnimationColorFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                AnimationColorFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                AnimationColorFileListRight.Add(AddWhiteSpace(1, "</states>"))
                            End If
                            StatesClosed = False
                            FirstStateFound = False
                            StateCount = 0
                        End If
                        If FirstConnectionFound Then
                            ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                            ' We do this at the end of the sub group iteration because we know by observation of the ME software
                            ' XML object creation that connections always go at the end
                            AnimationColorFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                            AnimationColorFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                            FirstConnectionFound = False
                        End If
                        CaptionCount = 0
                        ThresholdCount = 0

                        ' Close off this XML object
                        If OType = ECloseType.Complex Then
                            ' Requires complex object closure
                            AnimationColorFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                            AnimationColorFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                        End If
                        TestCount += 1
                        FirstCaptionFound = False
                    Next
                    ' This exit for is removed as we want to generate content for all animation color type objects found in the test
                    Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                AnimationColorFileListLeft.Add(itm)
                AnimationColorFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(AnimationColorFileListLeft)
            FormatXMLFile(AnimationColorFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(AnimationColorFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_animatecolor.xml"))
            WriteOutputFile(AnimationColorFileListRight, GetPathToLocalFile("Output", FnameVar & "R_animatecolor.xml"))
            WriteOutputFile(AnimationColorFileCSV, GetPathToLocalFile("Output", FnameVar & "animatecolor.csv"))



            MsgBox("")

        End If

#End Region

#Region "Active X Data List Generation"

        ' Check if this code block should run
        If MainObjectList.sList IsNot Nothing Then
            For Each itm As subparam In MainObjectList.sList.lSubParamList
                If itm.type = TypeConstants.Data Then
                    Type_ActiveXData_Exists = True
                End If
            Next
        End If


        If Type_ActiveXData_Exists Then
            ' Generate file content for the main parameter list
            Dim ActiveXDataFileListLeft As List(Of String) = New List(Of String)
            Dim ActiveXDataFileListRight As List(Of String) = New List(Of String)
            Dim ActiveXDataFileCSV As List(Of String) = New List(Of String)
            TestCount = 1
            FirstConnectionFound = False
            FirstCaptionFound = False ' Added to ensure only the first caption type gets processed when dealing with mutlistate objects
            FirstStateFound = False ' Reset the value here as it might still be set from the previous code block
            StateCount = 0
            CaptionCount = 0
            ThresholdCount = 0
            DataCount = 0
            Dim DataMask(10) As Boolean
            'Dim ThresholdInstCount As Integer = CountObjectInstance(MainObjectList, TypeConstants.Threshold)

            If MainObjectList.sList IsNot Nothing Then
                OType = ECloseType.Complex
            Else
                OType = ECloseType.Simple
            End If

            ' Initialise the left and right file lists with the header content
            For Each itm As String In HeaderList
                ActiveXDataFileListLeft.Add(itm)
                ActiveXDataFileListRight.Add(itm)
            Next

            ' Initialise the CSV test definition file list
            ActiveXDataFileCSV.Add("Test number,Property,Left,Right")

            '' Set up caption mask
            'Select Case ThresholdInstCount
            '    Case 1
            '        ThresholdMask(0) = True
            '        ThresholdMask(1) = False
            '    Case 2
            '        ThresholdMask(0) = True
            '        ThresholdMask(1) = True
            '    Case Else
            '        Throw New Exception("Whoops, it appears you didnt think of everything")
            'End Select

            ' This is active X Data so there will only be 1 data instance but we will maintain the previous framework for ease
            ' Of code comparison
            DataMask(0) = True
            DataMask(1) = False

            ' Loop through the test generation process for as many caption test masks are active
            For Dmask = 0 To 1
                If DataMask(Dmask) Then
                    For Each sublist As subparam In MainObjectList.sList.lSubParamList
                        If sublist.type = TypeConstants.Data Then
                            ' This object has a data sub object type so generate an output file for it
                            For Each sparam As Param In sublist.subParList
                                ' Generate the main objects data with only left cases
                                ActiveXDataFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                ActiveXDataFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
                                'ActiveXDataFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.caption, TestCount))

                                For Each subp As subparam In MainObjectList.sList.lSubParamList
                                    Select Case subp.type
                                        Case TypeConstants.Data
                                            If DataCount = Dmask Then
                                                ActiveXDataFileCSV.Add(CreateTestCaseByTestNumber(sparam, ValuePairList, ObjectTestClass.data, TestCount, ""))
                                                ActiveXDataFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                              sparam,
                                                                                              EditCase.Left,
                                                                                              1,
                                                                                              TestCount,
                                                                                              ValuePairList,
                                                                                              ObjectTestClass.data,
                                                                                              ECloseType.Simple))
                                                ActiveXDataFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                              sparam,
                                                                                              EditCase.Right,
                                                                                              1,
                                                                                              TestCount,
                                                                                              ValuePairList,
                                                                                              ObjectTestClass.data,
                                                                                              ECloseType.Simple))
                                            Else
                                                ActiveXDataFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.data,
                                                                                             ECloseType.Simple))
                                                ActiveXDataFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                                 EditCase.Left,
                                                                                                 1,
                                                                                                 TestCount,
                                                                                                 ValuePairList,
                                                                                                 ObjectTestClass.data,
                                                                                                 ECloseType.Simple))
                                            End If
                                            DataCount += 1

                                        Case TypeConstants.Threshold

                                            ActiveXDataFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))

                                            ActiveXDataFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.threshold,
                                                                                             ECloseType.Simple))

                                        Case TypeConstants.caption
                                            ActiveXDataFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                            ActiveXDataFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))

                                            CaptionCount += 1
                                        Case TypeConstants.imageSettings

                                            ActiveXDataFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))

                                            ActiveXDataFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             1,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.image,
                                                                                             ECloseType.Simple))
                                        Case TypeConstants.connection
                                            If FirstStateFound Then
                                                ' Close off the previous state before starting a connection block
                                                ActiveXDataFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                ActiveXDataFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                                ActiveXDataFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                                ActiveXDataFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                                StatesClosed = True
                                            End If
                                            If FirstConnectionFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                ActiveXDataFileListLeft.Add(AddWhiteSpace(1, "<connections>"))
                                                ActiveXDataFileListRight.Add(AddWhiteSpace(1, "<connections>"))
                                                FirstConnectionFound = True
                                            End If
                                            ActiveXDataFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                             EditCase.Left,
                                                                                             2,
                                                                                             TestCount,
                                                                                             ValuePairList,
                                                                                             ObjectTestClass.caption,
                                                                                             ECloseType.Simple))
                                            ActiveXDataFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                            EditCase.Left,
                                                                                            2,
                                                                                            TestCount,
                                                                                            ValuePairList,
                                                                                            ObjectTestClass.caption,
                                                                                            ECloseType.Simple))
                                        Case TypeConstants.state
                                            If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                                                ' if this is a subsequent state found after the first then close off the previous state
                                                ActiveXDataFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                                ActiveXDataFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                            End If
                                            If FirstStateFound = False Then
                                                ' Add an additional line here for the connection xml configuration on the first time only
                                                ActiveXDataFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                                                ActiveXDataFileListRight.Add(AddWhiteSpace(1, "<states>"))
                                                FirstStateFound = True
                                            End If
                                            ActiveXDataFileListLeft.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                         EditCase.Left,
                                                                                         2,
                                                                                         TestCount,
                                                                                         ValuePairList,
                                                                                         ObjectTestClass.state,
                                                                                         ECloseType.Complex))
                                            ActiveXDataFileListRight.Add(CreateXMLConnectionObjectByDefinition(subp,
                                                                                        EditCase.Left,
                                                                                        2,
                                                                                        TestCount,
                                                                                        ValuePairList,
                                                                                        ObjectTestClass.state,
                                                                                        ECloseType.Complex))
                                            StateCount += 1
                                        Case Else
                                            Throw New Exception("This type behaviour is not defined, please add it manually and try again")
                                    End Select

                                Next
                                If FirstStateFound Then
                                    If Not StatesClosed Then
                                        ' handle the case when no connection block is present and the state blocks need closed
                                        ActiveXDataFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                                        ActiveXDataFileListRight.Add(AddWhiteSpace(2, "</state>"))
                                        ActiveXDataFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                                        ActiveXDataFileListRight.Add(AddWhiteSpace(1, "</states>"))
                                    End If
                                    StatesClosed = False
                                    FirstStateFound = False
                                    StateCount = 0
                                End If
                                If FirstConnectionFound Then
                                    ' We know at least 1 connection has been defined and so we must close off the connections xml object group
                                    ' We do this at the end of the sub group iteration because we know by observation of the ME software
                                    ' XML object creation that connections always go at the end
                                    ActiveXDataFileListLeft.Add(AddWhiteSpace(1, "</connections>"))
                                    ActiveXDataFileListRight.Add(AddWhiteSpace(1, "</connections>"))
                                    FirstConnectionFound = False
                                End If
                                CaptionCount = 0
                                ThresholdCount = 0

                                ' Close off this XML object
                                If OType = ECloseType.Complex Then
                                    ' Requires complex object closure
                                    ActiveXDataFileListLeft.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                    ActiveXDataFileListRight.Add(AddWhiteSpace(0, "</" & MainObjectList.type & ">"))
                                End If
                                TestCount += 1
                                FirstCaptionFound = False
                            Next
                            Exit For ' added to avoid processing all captions when multiple instances exist as part of state sub objects
                        End If
                    Next
                End If
            Next



            ' Close off the files with the footer
            For Each itm As String In FooterList
                ActiveXDataFileListLeft.Add(itm)
                ActiveXDataFileListRight.Add(itm)
            Next

            'Format output file contents prior to writing
            FormatXMLFile(ActiveXDataFileListLeft)
            FormatXMLFile(ActiveXDataFileListRight)

            'Dim FnameVar As String = InputBox("Enter Output file name", "")

            WriteOutputFile(ActiveXDataFileListLeft, GetPathToLocalFile("Output", FnameVar & "L_data.xml"))
            WriteOutputFile(ActiveXDataFileListRight, GetPathToLocalFile("Output", FnameVar & "R_data.xml"))
            WriteOutputFile(ActiveXDataFileCSV, GetPathToLocalFile("Output", FnameVar & "data.csv"))



            MsgBox("")

        End If

#End Region




        MsgBox("Done")

    End Sub

    Public Function GetAnimationEcloseType(ByRef AnimationObjType As String) As ECloseType
        Select Case AnimationObjType
            Case "animateVisibility"
                Return ECloseType.Simple
            Case "animateColor"
                Return ECloseType.Complex
            Case "animateFill"
                Return ECloseType.Complex
            Case "animateHorizontalPosition"
                Return ECloseType.Complex
            Case "animateVerticalPosition"
                Return ECloseType.Complex
            Case "animateWidth"
                Return ECloseType.Complex
            Case "animateHeight"
                Return ECloseType.Complex
            Case "animateRotation"
                Return ECloseType.Complex
            Case Else
                Throw New Exception("Animation type not handled, please add manually and retry")
                Return ECloseType.Simple
        End Select
    End Function

    Public Function GetAddStrByAnimationType(ByRef AnimType As String) As String
        Select Case AnimType
            Case "animateVisibility"
                Return "Visibility-"
            Case "animateColor"
                Return "Color-"
            Case "animateFill"
                Return "Fill-"
            Case "color"
                Return "Animate Color Item:"
            Case "readFromTagExpressionRange"
                Return "Item "
            Case "constantExpressionRange"
                Return "Item "
            Case "defaultExpressionRange"
                Return ""
            Case "animateHorizontalPosition"
                Return "HorizontalPosition-"
            Case "animateVerticalPosition"
                Return "VerticalPosition-"
            Case "animateWidth"
                Return "Width-"
            Case "animateHeight"
                Return "Height-"
            Case "animateRotation"
                Return "Rotation-"
            Case Else
                Throw New Exception("Animation type not handled, please add manually and retry")
                Return ""
        End Select
    End Function

    Public Function CountObjectInstance(ByRef MainObjectList As ParamList, ByRef TConst As String) As Integer
        Dim i As Integer = 0
        Dim iSTR As String = TConst
        For Each subobj As subparam In MainObjectList.sList.lSubParamList.Where _
            (Function(x) x.type = iSTR)
            i += 1
        Next
        Return i
    End Function

    Public Function DetermineAddStrByCase(ByRef MainObjectList As ParamList,
                                          ByRef StateNo As Integer) As String
        Dim HasStates As Boolean
        For Each subobj As subparam In MainObjectList.sList.lSubParamList
            If subobj.type = TypeConstants.state Then
                HasStates = True
                Exit For
            End If
        Next

        Select Case HasStates
            Case True
                Return "State " & StateNo & " - "
            Case False
                Return ""
            Case Else
                Return ""
        End Select

    End Function

    Public Function GetTestCaseValueForConnectionBySubparam(ByRef MainObjectList As ParamList, ByRef ConnectionNo As Integer) As String

        Dim FilterCount As Integer = 1

        For Each itm As subparam In MainObjectList.sList.lSubParamList.Where _
            (Function(x) x.type = TypeConstants.connection)
            If FilterCount = ConnectionNo Then
                ' Return only this connection 
                Return itm.subParList.Item(1).sValue
            End If
            FilterCount += 1
        Next
        ' If the code gets here because the return value was not set return a default empty string
        Return ""

    End Function

    Public Sub AddConnectionsFilterParam(ByRef MainObjectList As ParamList,
                                          ByRef ConnectionFileList As List(Of String),
                                          ByRef ValuePairList As List(Of ValuePair),
                                          ByRef Filter As Integer)
        Dim FilterCount As Integer = 1

        ' Add the connections header
        ConnectionFileList.Add(AddWhiteSpace(1, "<connections>"))

        ' Loop through existing connections and add them all
        For Each itm As subparam In MainObjectList.sList.lSubParamList.Where _
            (Function(x) x.type = TypeConstants.connection)
            If Not FilterCount = Filter Then
                ConnectionFileList.Add(CreateXMLConnectionObjectByDefinition(itm,
                                                                         EditCase.Left,
                                                                         2,
                                                                         0,
                                                                         ValuePairList,
                                                                         ObjectTestClass.caption,
                                                                         ECloseType.Simple))
            Else
                ' Do nothing in here, this parameter is being deliberatly skipped to cause a difference
            End If

            FilterCount += 1

        Next

        ' Close the xml connection object
        ConnectionFileList.Add(AddWhiteSpace(1, "</connections>"))

    End Sub

    Public Sub AddConnectionsNoParams(ByRef MainObjectList As ParamList, ByRef ConnectionFileList As List(Of String), ByRef ValuePairList As List(Of ValuePair))

        ' Add the connections header
        ConnectionFileList.Add(AddWhiteSpace(1, "<connections>"))

        ' Loop through existing connections and add them all
        For Each itm As subparam In MainObjectList.sList.lSubParamList.Where _
            (Function(x) x.type = TypeConstants.connection)
            ConnectionFileList.Add(CreateXMLConnectionObjectByDefinition(itm,
                                                                         EditCase.Left,
                                                                         2,
                                                                         0,
                                                                         ValuePairList,
                                                                         ObjectTestClass.caption,
                                                                         ECloseType.Simple))
        Next

        ' Close the xml connection object
        ConnectionFileList.Add(AddWhiteSpace(1, "</connections>"))

    End Sub

    Public Sub GenerateXMLObjectWithoutConnections(ByRef MainObjectList As ParamList,
                                                   ByRef ConnectionFileListLeft As List(Of String),
                                                   ByRef ConnectionFileListRight As List(Of String),
                                                   ByRef ConnectionFileCSV As List(Of String),
                                                   ByRef TestCount As Integer,
                                                   ByRef ValuePairList As List(Of ValuePair),
                                                   ByRef FirstConnectionFound As Boolean,
                                                   ByRef OType As ECloseType)

        Dim FirstStateFound As Boolean = False
        Dim StatesClosed As Boolean = False
        Dim FirstAnimationFound As Boolean = False
        Dim AnimationsClosed As Boolean = False

        ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))
        ConnectionFileListRight.Add(CreateXMLObjectByDefinition(MainObjectList, MainObjectList.pList.Item(1), EditCase.Left, 0, TestCount, ValuePairList, "", OType))

        For Each subp As subparam In MainObjectList.sList.lSubParamList
            Select Case subp.type
                Case TypeConstants.Animate
                    If FirstAnimationFound = False Then
                        ' Add an additional line here for the connection xml configuration on the first time only
                        ConnectionFileListLeft.Add(AddWhiteSpace(1, "<animations>"))
                        ConnectionFileListRight.Add(AddWhiteSpace(1, "<animations>"))
                        FirstAnimationFound = True
                    End If
                    ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.animate,
                                                                                     ECloseType.Simple))
                    ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.animate,
                                                                                     ECloseType.Simple))
                Case TypeConstants.caption
                    ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.caption,
                                                                                     ECloseType.Simple))
                    ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.caption,
                                                                                     ECloseType.Simple))


                Case TypeConstants.imageSettings

                    ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.image,
                                                                                     ECloseType.Simple))

                    ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.image,
                                                                                     ECloseType.Simple))
                Case TypeConstants.connection
                    ' Close any open states and do not process the connection
                    If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                        ' if this is a subsequent state found after the first then close off the previous state
                        If Not StatesClosed Then
                            ConnectionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                            ConnectionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                            ConnectionFileListLeft.Add(AddWhiteSpace(1, "</states>"))
                            ConnectionFileListRight.Add(AddWhiteSpace(1, "</states>"))
                            StatesClosed = True
                        End If
                    End If
                    If FirstAnimationFound Then
                        If Not AnimationsClosed Then
                            ConnectionFileListLeft.Add(AddWhiteSpace(2, "</animations>"))
                            ConnectionFileListRight.Add(AddWhiteSpace(2, "</animations>"))
                        End If
                    End If
                        Case TypeConstants.state
                    If FirstStateFound Then ' deliberately placed before the first state found detector so it will only trigger on subsequent states
                        ' if this is a subsequent state found after the first then close off the previous state
                        ConnectionFileListLeft.Add(AddWhiteSpace(2, "</state>"))
                        ConnectionFileListRight.Add(AddWhiteSpace(2, "</state>"))
                    End If
                    If FirstStateFound = False Then
                        ' Add an additional line here for the connection xml configuration on the first time only
                        ConnectionFileListLeft.Add(AddWhiteSpace(1, "<states>"))
                        ConnectionFileListRight.Add(AddWhiteSpace(1, "<states>"))
                        FirstStateFound = True
                    End If
                    ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.state,
                                                                                     ECloseType.Complex))

                    ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.state,
                                                                                     ECloseType.Complex))
                Case TypeConstants.Threshold
                    ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.threshold,
                                                                                     ECloseType.Simple))

                    ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.threshold,
                                                                                     ECloseType.Simple))
                Case TypeConstants.Data
                    ConnectionFileListLeft.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.data,
                                                                                     ECloseType.Simple))

                    ConnectionFileListRight.Add(CreateXMLObjectByDefinition(subp,
                                                                                     EditCase.Left,
                                                                                     1,
                                                                                     TestCount,
                                                                                     ValuePairList,
                                                                                     ObjectTestClass.data,
                                                                                     ECloseType.Simple))

                Case Else
                    Throw New Exception("This type behaviour is not defined, please add it manually and try again")
            End Select
        Next

    End Sub

    Public Sub FormatXMLFile(ByRef Flist As List(Of String))
        For a = 0 To Flist.Count - 1
            Flist.Item(a) = Flist.Item(a).Replace("=""True", "=""true")
            Flist.Item(a) = Flist.Item(a).Replace("=""False", "=""false")
        Next
    End Sub

    Public Sub WriteOutputFile(ByRef Flist As List(Of String), ByRef Fpath As String)
        Using writer As StreamWriter = New StreamWriter(Fpath, False)
            For Each itm As String In Flist
                writer.WriteLine(itm)
            Next
        End Using
    End Sub

    Public Function CreateTestCaseByConnectionSpecial(ByRef msg As String,
                                                      ByRef lval As String,
                                                      ByRef rval As String,
                                                      ByRef Tcase As Integer) As String
        Dim tstr As String = ""
        Dim LparVal As String = lval
        Dim RparVal As String = rval
        Dim ParName As String = msg
        tstr = Tcase.ToString & "," & ParName & "," & LparVal & "," & RparVal
        CreateTestCaseByConnectionSpecial = tstr

    End Function

    Public Function CreateTestCaseByConnection(ByRef Par As Param,
                                               ByRef Tcase As Integer,
                                               ByRef ConnectionNo As Integer) As String
        Dim tstr As String = ""
        Dim LparVal As String = Par.sValue
        Dim RparVal As String = Par.sValue & "s"
        Dim ParName As String = "Connection " & ConnectionNo & " - " & Par.sProperty
        tstr = Tcase.ToString & "," & ParName & "," & LparVal & "," & RparVal
        CreateTestCaseByConnection = tstr

    End Function

    Public Function CreateTestCaseByTestNumber(ByRef Par As Param,
                                            ByRef Tlist As List(Of ValuePair),
                                            ByRef oClass As String,
                                            ByRef Tcase As Integer) As String
        Dim tstr As String = ""
        Dim LparVal As String = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Left)
        Dim RparVal As String = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Right)
        Dim ParName As String = Par.sProperty
        tstr = Tcase.ToString & "," & ParName & "," & LparVal & "," & RparVal
        CreateTestCaseByTestNumber = tstr

    End Function

    Public Function CreateTestCaseByTestNumber(ByRef Par As Param,
                                               ByRef Tlist As List(Of ValuePair),
                                               ByRef oClass As String,
                                               ByRef Tcase As Integer,
                                               ByRef AddDescription As String) As String
        Dim AlignmentType As String = ""
        Dim tstr As String = ""
        Dim LparVal As String
        Dim RparVal As String
        Dim ParName As String = Par.sProperty

        ' Handle special cases of test values for certain edge cases
        Select Case Par.sProperty
            Case "alignment"
                Select Case Par.sValue
                    Case "left"
                        AlignmentType = "simple"
                    Case "center"
                        AlignmentType = "simple"
                    Case "right"
                        AlignmentType = "simple"
                    Case Else
                        AlignmentType = "complex"
                End Select

                If AlignmentType = "simple" Then
                    LparVal = "center"
                    RparVal = "left"
                Else
                    LparVal = "middleCenter"
                    RparVal = "middleLeft"
                End If
            Case "value"
                LparVal = "0" ' Use original value from object
                RparVal = "-1" ' fixed value from the other special case handler
            Case "thresholdIndex"
                ' retain original values
                LparVal = Par.sValue
                RparVal = Par.sValue
            Case Else
                ' Gather test values in the normal fashion 
                LparVal = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Left)
                RparVal = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Right)
        End Select
        ' Handle special case of alignment property and value
        If Par.sProperty = "alignment" Then



        Else

        End If

        tstr = Tcase.ToString & "," & AddDescription & ParName & "," & LparVal & "," & RparVal
        CreateTestCaseByTestNumber = tstr

    End Function

    Public Function CreateTestCaseByTestNumber(ByRef Par As Param,
                                               ByRef Tlist As List(Of ValuePair),
                                               ByRef oClass As String,
                                               ByRef Tcase As Integer,
                                               ByRef AddDescription As String,
                                               ByRef ParentType As String) As String
        Dim AlignmentType As String = ""
        Dim tstr As String = ""
        Dim LparVal As String
        Dim RparVal As String
        Dim ParName As String = Par.sProperty

        ' Handle special cases of test values for certain edge cases
        Select Case Par.sProperty
            Case "alignment"
                Select Case Par.sValue
                    Case "left"
                        AlignmentType = "simple"
                    Case "center"
                        AlignmentType = "simple"
                    Case "right"
                        AlignmentType = "simple"
                    Case Else
                        AlignmentType = "complex"
                End Select

                If AlignmentType = "simple" Then
                    LparVal = "center"
                    RparVal = "left"
                Else
                    LparVal = "middleCenter"
                    RparVal = "middleLeft"
                End If
            Case "value"
                If Not ParentType = TypeConstants.Color Then
                    LparVal = "0" ' Use original value from object
                    RparVal = "-1" ' fixed value from the other special case handler
                Else
                    LparVal = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Left)
                    RparVal = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Right)
                End If

            Case "thresholdIndex"
                ' retain original values
                LparVal = Par.sValue
                RparVal = Par.sValue
            Case Else
                ' Gather test values in the normal fashion 
                LparVal = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Left)
                RparVal = GetParameterValueByCase(Par, Tlist, oClass, EditCase.Right)
        End Select
        ' Handle special case of alignment property and value
        If Par.sProperty = "alignment" Then



        Else

        End If

        tstr = Tcase.ToString & "," & AddDescription & ParName & "," & LparVal & "," & RparVal
        CreateTestCaseByTestNumber = tstr

    End Function

    Public Function CreateXMLObjectByDefinition(ByRef parList As ParamList,
                                                ByRef EditParam As Param,
                                                ByRef ECase As EditCase,
                                                ByRef IndentLevel As Integer,
                                                ByRef TestInst As Integer,
                                                ByRef TestDefs As List(Of ValuePair),
                                                ByRef TestClass As String,
                                                ByRef ClosingType As ECloseType) As String
        ' Create an entire instance of an xml object on a single line based on the supplied definition in the param list
        ' Substitute 1 of the parameters with the left/right special case based on matching the current edit param
        ' All other parameters recieve the default values

        Dim tstr As String = "" ' creating a temporary string here because the name is shorter than the function name for brevity

        ' Get the special case parameters for this object
        Dim EditCaseVal As String = GetParameterValueByCase(EditParam, TestDefs, TestClass, ECase)
        Dim AlignmentType As String = ""

        ' Begin creating the xml object string
        tstr &= "<" & parList.type & " "
        tstr &= "name=""" & parList.type & "_Test" & TestInst & """ "

        ' Loop through each property of the XML object and create an entry for it in the XML string
        ' Find the 1 property whose value needs to be changed and substitute its matching right value from the valuepair list
        For Each itm In parList.pList
            Select Case itm.sProperty
                Case "name"
                    ' do nothing, dont process this because the name field is special
                Case "thresholdIndex"
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                Case "alignment"
                    ' It appears rockwell were naughty and use 2 xml properties called alignment
                    ' Where the type of object might need them as 1 of 2 types of enum on deserialization
                    ' This block will try to read theh type of alignment attribute required and select the correct type
                    ' of test case to generate
                    Select Case itm.sValue
                        Case "left"
                            AlignmentType = "simple"
                        Case "center"
                            AlignmentType = "simple"
                        Case "right"
                            AlignmentType = "simple"
                        Case Else
                            AlignmentType = "complex"
                    End Select
                    If AlignmentType = "simple" Then
                        If itm.sProperty = EditParam.sProperty Then
                            If ECase = EditCase.Left Then
                                tstr &= itm.sProperty & "=""" & "center" & """ "
                            Else
                                tstr &= itm.sProperty & "=""" & "left" & """ "
                            End If
                        Else
                            tstr &= itm.sProperty & "=""" & "center" & """ "
                        End If
                    Else
                        If itm.sProperty = EditParam.sProperty Then
                            If ECase = EditCase.Left Then
                                tstr &= itm.sProperty & "=""" & "middleCenter" & """ "
                            Else
                                tstr &= itm.sProperty & "=""" & "middleLeft" & """ "
                            End If
                        Else
                            tstr &= itm.sProperty & "=""" & "middleCenter" & """ "
                        End If
                    End If
                Case Else
                    If itm.sProperty = EditParam.sProperty Then
                        tstr &= itm.sProperty & "=""" & EditCaseVal & """ "
                    Else
                        tstr &= itm.sProperty & "=""" & GetParameterValueByCase(itm, TestDefs, TestClass, EditCase.Left) & """ "
                    End If
            End Select
        Next

        ' add the closure of the xml object depending on type
        Select Case ClosingType
            Case ECloseType.Simple
                tstr &= "/>"
            Case ECloseType.Complex
                tstr &= ">"
        End Select

        ' add the required whitespace 
        tstr = AddWhiteSpace(IndentLevel, tstr)

        ' Finally return the completed xml object
        CreateXMLObjectByDefinition = tstr


    End Function

    Public Function CreateXMLObjectByDefinition(ByRef parList As subparam,
                                                ByRef EditParam As Param,
                                                ByRef ECase As EditCase,
                                                ByRef IndentLevel As Integer,
                                                ByRef TestInst As Integer,
                                                ByRef TestDefs As List(Of ValuePair),
                                                ByRef TestClass As String,
                                                ByRef ClosingType As ECloseType) As String
        ' Create an entire instance of an xml object on a single line based on the supplied definition in the param list
        ' Substitute 1 of the parameters with the left/right special case based on matching the current edit param
        ' All other parameters recieve the default values

        Dim tstr As String = "" ' creating a temporary string here because the name is shorter than the function name for brevity
        Dim AlignmentType As String = ""

        ' Get the special case parameters for this object
        Dim EditCaseVal As String = GetParameterValueByCase(EditParam, TestDefs, TestClass, ECase)

        ' Begin creating the xml object string
        tstr &= "<" & parList.type & " "

        Dim CheckAnimationType As Boolean
        Select Case True
            Case parList.type Like "animate*"
                CheckAnimationType = True
        End Select

        If Not CheckAnimationType Then
            tstr &= "name=""" & parList.type & "_Test" & TestInst & """ "
        End If


        ' Loop through each property of the XML object and create an entry for it in the XML string
        ' Find the 1 property whose value needs to be changed and substitute its matching right value from the valuepair list
        For Each itm In parList.subParList
            Select Case itm.sProperty
                Case "name"
                    ' do nothing, dont process this because the name field is special
                Case "thresholdIndex"
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                Case "alignment"
                    ' It appears rockwell were naughty and use 2 xml properties called alignment
                    ' Where the type of object might need them as 1 of 2 types of enum on deserialization
                    ' This block will try to read theh type of alignment attribute required and select the correct type
                    ' of test case to generate
                    Select Case itm.sValue
                        Case "left"
                            AlignmentType = "simple"
                        Case "center"
                            AlignmentType = "simple"
                        Case "right"
                            AlignmentType = "simple"
                        Case Else
                            AlignmentType = "complex"
                    End Select
                    If AlignmentType = "simple" Then
                        If itm.sProperty = EditParam.sProperty Then
                            If ECase = EditCase.Left Then
                                tstr &= itm.sProperty & "=""" & "center" & """ "
                            Else
                                tstr &= itm.sProperty & "=""" & "left" & """ "
                            End If
                        Else
                            tstr &= itm.sProperty & "=""" & "center" & """ "
                        End If
                    Else
                        If itm.sProperty = EditParam.sProperty Then
                            If ECase = EditCase.Left Then
                                tstr &= itm.sProperty & "=""" & "middleCenter" & """ "
                            Else
                                tstr &= itm.sProperty & "=""" & "middleLeft" & """ "
                            End If
                        Else
                            tstr &= itm.sProperty & "=""" & "middleCenter" & """ "
                        End If
                    End If
                Case Else
                    If itm.sProperty = EditParam.sProperty Then
                        tstr &= itm.sProperty & "=""" & EditCaseVal & """ "
                    Else
                        tstr &= itm.sProperty & "=""" & GetParameterValueByCase(itm, TestDefs, TestClass, EditCase.Left) & """ "
                    End If
            End Select
        Next

        ' add the closure of the xml object depending on type
        Select Case ClosingType
            Case ECloseType.Simple
                tstr &= "/>"
            Case ECloseType.Complex
                tstr &= ">"
        End Select

        ' add the required whitespace 
        tstr = AddWhiteSpace(IndentLevel, tstr)

        ' Finally return the completed xml object
        CreateXMLObjectByDefinition = tstr


    End Function

    Public Function CreateXMLObjectStateByDefinition(ByRef parList As subparam,
                                                ByRef EditParam As Param,
                                                ByRef ECase As EditCase,
                                                ByRef IndentLevel As Integer,
                                                ByRef TestInst As Integer,
                                                ByRef TestDefs As List(Of ValuePair),
                                                ByRef TestClass As String,
                                                ByRef ClosingType As ECloseType) As String
        ' Create an entire instance of an xml object on a single line based on the supplied definition in the param list
        ' Substitute 1 of the parameters with the left/right special case based on matching the current edit param
        ' All other parameters recieve the default values

        Dim tstr As String = "" ' creating a temporary string here because the name is shorter than the function name for brevity

        ' Get the special case parameters for this object
        Dim EditCaseVal As String = GetParameterValueByCase(EditParam, TestDefs, TestClass, ECase)

        ' Begin creating the xml object string
        tstr &= "<" & parList.type & " "

        ' Loop through each property of the XML object and create an entry for it in the XML string
        ' Find the 1 property whose value needs to be changed and substitute its matching right value from the valuepair list
        For Each itm In parList.subParList
            Select Case itm.sProperty
                Case "name"
                    ' do nothing, dont process this because the name field is special
                Case "stateId"
                    ' Preserve the original value of this object parameter
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                Case "value"
                    ' normally we want to preserve the original value of this property
                    ' If the value property is the current selector, then use the original value for left, and then use -1 for the right
                    If EditParam.sProperty = "value" Then
                        ' left use original, right use - 1
                        Select Case ECase
                            Case EditCase.Left
                                tstr &= itm.sProperty & "=""" & "0" & """ "
                            Case EditCase.Right
                                tstr &= itm.sProperty & "=""" & "-1" & """ "
                        End Select
                    Else
                        ' use original for both sides 
                        tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                    End If
                Case Else
                    If itm.sProperty = EditParam.sProperty Then
                        tstr &= itm.sProperty & "=""" & EditCaseVal & """ "
                    Else
                        tstr &= itm.sProperty & "=""" & GetParameterValueByCase(itm, TestDefs, TestClass, EditCase.Left) & """ "
                    End If
            End Select
        Next

        ' add the closure of the xml object depending on type
        Select Case ClosingType
            Case ECloseType.Simple
                tstr &= "/>"
            Case ECloseType.Complex
                tstr &= ">"
        End Select

        ' add the required whitespace 
        tstr = AddWhiteSpace(IndentLevel, tstr)

        ' Finally return the completed xml object
        Return tstr

    End Function
    Public Function CreateXMLObjectByDefinition(ByRef parList As subparam,
                                                ByRef ECase As EditCase,
                                                ByRef IndentLevel As Integer,
                                                ByRef TestInst As Integer,
                                                ByRef TestDefs As List(Of ValuePair),
                                                ByRef TestClass As String,
                                                ByRef ClosingType As ECloseType) As String
        ' Create an entire instance of an xml object on a single line based on the supplied definition in the param list
        ' Substitute 1 of the parameters with the left/right special case based on matching the current edit param
        ' All other parameters recieve the default values

        Dim tstr As String = "" ' creating a temporary string here because the name is shorter than the function name for brevity
        Dim AlignmentType As String = ""

        ' Get the special case parameters for this object
        'Dim EditCaseVal As String = GetParameterValueByCase(EditParam, TestDefs, TestClass, ECase)

        ' Begin creating the xml object string
        tstr &= "<" & parList.type & " "
        'tstr &= "name=""" & parList.type & "_Test" & TestInst & """ "

        ' Loop through each property of the XML object and create an entry for it in the XML string
        ' Find the 1 property whose value needs to be changed and substitute its matching right value from the valuepair list
        For Each itm In parList.subParList
            Select Case itm.sProperty
                'Case "name"
                ' do nothing, dont process this because the name field is special
                Case "thresholdIndex"
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                Case "alignment"
                    ' It appears rockwell were naughty and use 2 xml properties called alignment
                    ' Where the type of object might need them as 1 of 2 types of enum on deserialization
                    ' This block will try to read theh type of alignment attribute required and select the correct type
                    ' of test case to generate
                    Select Case itm.sValue
                        Case "left"
                            AlignmentType = "simple"
                        Case "center"
                            AlignmentType = "simple"
                        Case "right"
                            AlignmentType = "simple"
                        Case Else
                            AlignmentType = "complex"
                    End Select
                    If AlignmentType = "simple" Then
                        tstr &= itm.sProperty & "=""" & "center" & """ "
                        'If itm.sProperty = EditParam.sProperty Then
                        '    If ECase = EditCase.Left Then
                        '        tstr &= itm.sProperty & "=""" & "center" & """ "
                        '    Else
                        '        tstr &= itm.sProperty & "=""" & "left" & """ "
                        '    End If
                        'Else

                        'End If
                    Else
                        tstr &= itm.sProperty & "=""" & "middleCenter" & """ "
                        'If itm.sProperty = EditParam.sProperty Then
                        '    If ECase = EditCase.Left Then
                        '        tstr &= itm.sProperty & "=""" & "middleCenter" & """ "
                        '    Else
                        '        tstr &= itm.sProperty & "=""" & "middleLeft" & """ "
                        '    End If
                        'Else

                        'End If
                    End If
                Case Else
                    'If itm.sProperty = EditParam.sProperty Then
                    '    tstr &= itm.sProperty & "=""" & EditCaseVal & """ "
                    'Else
                    '    tstr &= itm.sProperty & "=""" & GetParameterValueByCase(itm, TestDefs, TestClass, EditCase.Left) & """ "
                    'End If
                    tstr &= itm.sProperty & "=""" & GetParameterValueByCase(itm, TestDefs, TestClass, EditCase.Left) & """ "
            End Select
        Next

        ' add the closure of the xml object depending on type
        Select Case ClosingType
            Case ECloseType.Simple
                tstr &= "/>"
            Case ECloseType.Complex
                tstr &= ">"
        End Select

        ' add the required whitespace 
        tstr = AddWhiteSpace(IndentLevel, tstr)

        ' Finally return the completed xml object
        CreateXMLObjectByDefinition = tstr


    End Function

    Public Function CreateXMLConnectionObjectByDefinition(ByRef parList As subparam,
                                                ByRef ECase As EditCase,
                                                ByRef IndentLevel As Integer,
                                                ByRef TestInst As Integer,
                                                ByRef TestDefs As List(Of ValuePair),
                                                ByRef TestClass As String,
                                                ByRef ClosingType As ECloseType) As String
        ' Create an entire instance of an xml object on a single line based on the supplied definition in the param list
        ' Substitute 1 of the parameters with the left/right special case based on matching the current edit param
        ' All other parameters recieve the default values

        Dim tstr As String = "" ' creating a temporary string here because the name is shorter than the function name for brevity

        ' Get the special case parameters for this object
        'Dim EditCaseVal As String = GetParameterValueByCase(EditParam, TestDefs, TestClass, ECase)

        ' Begin creating the xml object string
        tstr &= "<" & parList.type & " "
        'tstr &= "name=""" & parList.type & "_Test" & TestInst & """ "

        ' Loop through each property of the XML object and create an entry for it in the XML string
        ' Find the 1 property whose value needs to be changed and substitute its matching right value from the valuepair list
        For Each itm In parList.subParList
            Select Case itm.sProperty
                Case "name"
                    ' use the parent object name here because the object name is not subject to test as it is a fixed string from the ME application
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                Case "expression"
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                Case Else
                    'If itm.sProperty = EditParam.sProperty Then
                    '    tstr &= itm.sProperty & "=""" & EditCaseVal & """ "
                    'Else
                    '    tstr &= itm.sProperty & "=""" & GetParameterValueByCase(itm, TestDefs, TestClass, EditCase.Left) & """ "
                    'End If
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
            End Select
        Next

        ' add the closure of the xml object depending on type
        Select Case ClosingType
            Case ECloseType.Simple
                tstr &= "/>"
            Case ECloseType.Complex
                tstr &= ">"
        End Select

        ' add the required whitespace 
        tstr = AddWhiteSpace(IndentLevel, tstr)

        ' Finally return the completed xml object
        CreateXMLConnectionObjectByDefinition = tstr


    End Function

    Public Function CreateTestXMLConnectionObjectByDefinition(ByRef parList As subparam,
                                                          ByRef ECase As EditCase,
                                                          ByRef IndentLevel As Integer,
                                                          ByRef TestInst As Integer,
                                                          ByRef TestDefs As List(Of ValuePair),
                                                          ByRef TestClass As String,
                                                          ByRef ClosingType As ECloseType,
                                                          ByRef TestChar As String) As String
        ' Create an entire instance of an xml object on a single line based on the supplied definition in the param list
        ' Substitute 1 of the parameters with the left/right special case based on matching the current edit param
        ' All other parameters recieve the default values

        Dim tstr As String = "" ' creating a temporary string here because the name is shorter than the function name for brevity

        ' Get the special case parameters for this object
        'Dim EditCaseVal As String = GetParameterValueByCase(EditParam, TestDefs, TestClass, ECase)

        ' Begin creating the xml object string
        tstr &= "<" & parList.type & " "
        'tstr &= "name=""" & parList.type & "_Test" & TestInst & """ "

        ' Loop through each property of the XML object and create an entry for it in the XML string
        ' Find the 1 property whose value needs to be changed and substitute its matching right value from the valuepair list
        For Each itm In parList.subParList
            Select Case itm.sProperty
                Case "name"
                    ' use the parent object name here because the object name is not subject to test as it is a fixed string from the ME application
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
                Case "expression"
                    tstr &= itm.sProperty & "=""" & itm.sValue & TestChar & """ "
                Case Else
                    'If itm.sProperty = EditParam.sProperty Then
                    '    tstr &= itm.sProperty & "=""" & EditCaseVal & """ "
                    'Else
                    '    tstr &= itm.sProperty & "=""" & GetParameterValueByCase(itm, TestDefs, TestClass, EditCase.Left) & """ "
                    'End If
                    tstr &= itm.sProperty & "=""" & itm.sValue & """ "
            End Select
        Next

        ' add the closure of the xml object depending on type
        Select Case ClosingType
            Case ECloseType.Simple
                tstr &= "/>"
            Case ECloseType.Complex
                tstr &= ">"
        End Select

        ' add the required whitespace 
        tstr = AddWhiteSpace(IndentLevel, tstr)

        ' Finally return the completed xml object
        CreateTestXMLConnectionObjectByDefinition = tstr


    End Function
    Public Function RemoveSpacesAndFlatten(ByRef str As String) As String
        RemoveSpacesAndFlatten = str.Replace(" ", "").ToLower
    End Function

    Public Function GetParameterValueByCase(ByRef Par As Param,
                                            ByRef Tlist As List(Of ValuePair),
                                            ByRef oClass As String,
                                            ByRef EditCase As EditCase) As String
        GetParameterValueByCase = ""
        For Each itm In Tlist
            If itm.name = Par.sProperty Then
                If itm.oClass = oClass Then
                    If EditCase = EditCase.Left Then GetParameterValueByCase = itm.Value1
                    If EditCase = EditCase.Right Then GetParameterValueByCase = itm.Value2
                End If
            End If
        Next

    End Function

    ''' <summary>
    ''' Adds whitespace to a string to a specified indent level, supports zero indents
    ''' </summary>
    ''' <param name="indentLevel">specify the indent level to add to the strings whitespace</param>
    ''' <param name="inputString">input string to add the whitespace to</param>
    ''' <returns>the input string with the whitespace added at the begining of the string</returns>
    Public Function AddWhiteSpace(ByVal indentLevel As Integer, ByRef inputString As String) As String
        AddWhiteSpace = "    " ' add the first 4 lines of whitespace as default
        For a = 0 To indentLevel
            AddWhiteSpace &= "    "
        Next
        AddWhiteSpace &= inputString
    End Function

    Public Function ReadFile(ByRef filelist As List(Of String), ByRef filepath As String) As List(Of String)
        ReadFile = New List(Of String)
        Using reader As New StreamReader(filepath)
            Do
                ReadFile.Add(reader.ReadLine)
            Loop Until reader.EndOfStream
        End Using

    End Function
    Public Sub Outputreport(ByRef StringList As List(Of String))
        Using output As StreamWriter = New StreamWriter("C:\temp\TestFiles\Output.txt", False)
            For a = 0 To StringList.Count - 1
                output.WriteLine(StringList.Item(a))
            Next
        End Using
        System.Diagnostics.Process.Start("notepad.exe", "C:\temp\TestFiles\Output.txt")
    End Sub

    Function ReturnFormattedValues(ByRef strval As String) As String
        Select Case strval
            Case "TRUE"
                ReturnFormattedValues = "True"
            Case "FALSE"
                ReturnFormattedValues = "False"
            Case Else
                ReturnFormattedValues = strval ' no formating required so return the original value
        End Select
        If InStr(strval, " ") Then
            'Throw New Exception("Value pairs with spaces cannot be accepted")
        End If
    End Function

    Public Function GetSubObjParams(ByRef strlne As String) As subparam
        Dim tarr()
        Dim SkipCount As Integer = 0
        tarr = Split(strlne, " ")
        GetSubObjParams = New subparam

        ' create a params object containing all the property and value pairs
        For a = 0 To tarr.Length - 1
            If Not tarr(a) = "" Then ' avoid empty array elements, these are just converted white space
                If InStr(tarr(a), "=") = 0 Then
                    ' you found the header, extract it and add it to the listobject
                    If GetSubObjParams.Headerset = False Then
                        Dim hdtarr()
                        hdtarr = Split(tarr(a), "<")
                        GetSubObjParams.type = hdtarr(1)
                        GetSubObjParams.Headerset = True
                    End If
                    ' still need to account for skipcount integer here because the param that is split will not have an "=" symbol in its char string
                    If Not SkipCount = 0 Then
                        SkipCount -= 1
                    End If
                Else
                    If Not SkipCount = 0 Then
                        SkipCount -= 1
                    Else
                        Select Case CountCharsInStringOccurence(tarr(a))
                            Case 1
                                ' This parameter is split over more than 1 array location as its name has a space in it
                                ' Find the ending of the parameter and reconstitute it
                                ' Mark the function to skip the next param
                                Dim tstr As String = tarr(a)
                                For b = (a + 1) To tarr.Length - 1
                                    ' search the remainder of the array for the ending of the current param starting from the next position
                                    tstr &= " " & tarr(b)
                                    If CountCharsInStringOccurence(tarr(b)) = 1 Then
                                        ' found the ending
                                        ' work out how many array locations to skip
                                        SkipCount = b - a
                                        Dim tObj As Param = New Param
                                        tObj = ExtractParamObject(tstr)
                                        If tObj IsNot Nothing Then
                                            GetSubObjParams.subParList.Add(tObj)
                                        End If
                                        Exit For
                                    End If
                                Next
                            Case 2
                                ' Normal case when the param has been extracted in full
                                Dim tObj As Param = New Param
                                tObj = ExtractParamObject(tarr(a))
                                If tObj IsNot Nothing Then
                                    GetSubObjParams.subParList.Add(tObj)
                                End If
                            Case Else
                                ' I dont expect this to happen but here it is
                                Throw New Exception("What the swear word?")

                        End Select
                    End If


                End If
            End If
        Next
    End Function

    Public Function CountCharsInStringOccurence(ByRef str As String) As Integer
        Dim count As Integer = 0
        For Each c As Char In str
            If c = """" Then
                count += 1
            End If
        Next
        CountCharsInStringOccurence = count
    End Function

    Public Function GetObjParams(ByRef strlne As String) As ParamList
        Dim tarr()
        tarr = Split(strlne, " ")
        GetObjParams = New ParamList

        ' create a params object containing all the property and value pairs
        For a = 0 To tarr.Length - 1
            If Not tarr(a) = "" Then ' avoid empty array elements, these are just converted white space
                If InStr(tarr(a), "=") = 0 Then
                    ' you found the header, extract it and add it to the listobject
                    If GetObjParams.Headerset = False Then
                        Dim hdtarr()
                        hdtarr = Split(tarr(a), "<")
                        GetObjParams.type = hdtarr(1)
                        GetObjParams.Headerset = True
                    End If
                Else
                    Dim tObj As Param = New Param
                    tObj = ExtractParamObject(tarr(a))
                    If tObj IsNot Nothing Then
                        GetObjParams.pList.Add(tObj)
                    End If
                End If
            End If
        Next
    End Function

    Public Function ExtractParamObject(ByVal ArrayElement As String) As Param
        ExtractParamObject = New Param
        Dim tarr()
        tarr = Split(ArrayElement, "=")
        If tarr.Length > 1 Then
            ' create an object from it
            ExtractParamObject.sProperty = tarr(0)
            ExtractParamObject.sValue = RemoveDoubleQuotes(tarr(1))

            'MsgBox(tarr(0))
        Else
            ' return nothing because it didnt fit the use case
            ExtractParamObject = Nothing
        End If
    End Function

    Private Function RemoveDoubleQuotes(ByRef str As String)
        Dim tarr()
        tarr = Split(str, """")
        RemoveDoubleQuotes = tarr(1) ' return inner content
    End Function

    Public Function GetPathToDefFile(ByVal Name As String) As String

        ' Build the path string
        Dim wrkstr As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Test Definitions\"
        wrkstr = wrkstr & Name
        wrkstr = wrkstr & ".csv"
        GetPathToDefFile = wrkstr

    End Function

    ''' <summary>
    ''' Returns a fully qualified path to a folder/file combination within the current executing assembly location
    ''' </summary>
    ''' <param name="Folder">Name of the folder in the assembly path, can handle multiple folder layers if required</param>
    ''' <param name="File">Name of the file plus the extension</param>
    ''' <returns>Returns a fully qualified path to a folder/file combination within the current executing assembly location</returns>
    Public Function GetPathToLocalFile(ByVal Folder As String, ByVal File As String) As String
        GetPathToLocalFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\" & Folder & "\"
        GetPathToLocalFile = GetPathToLocalFile & File
    End Function

End Class

Public Class ParamList
    Public type As String
    Public Headerset As Boolean
    Public pList As List(Of Param) ' list of main parameters
    Public sList As SubParamList ' object with lists of sub object parameters if required
    Public Sub New()
        pList = New List(Of Param)
        Headerset = False
    End Sub
End Class
Public Class Param
    Public sProperty As String
    Public sValue As String
End Class

Public Class SubParamList
    Public lSubParamList As List(Of subparam)
    Public Sub New()
        lSubParamList = New List(Of subparam)
    End Sub
End Class

Public Class subparam
    Public type As String
    Public Headerset As Boolean
    Public subParList As List(Of Param)
    Public Sub New()
        subParList = New List(Of Param)
    End Sub
End Class

Public Class ValuePair
    Public oClass As String
    Public name As String
    Public Value1 As String
    Public Value2 As String
    Public Sub New(ByRef o As String,
                   ByRef n As String,
                   ByRef v1 As String,
                   ByRef v2 As String)
        oClass = o
        name = n
        Value1 = v1
        Value2 = v2
    End Sub
End Class

Public Class TypeConstants
    Public Const caption As String = "caption"
    Public Const imageSettings As String = "imageSettings"
    Public Const connections As String = "connections"
    Public Const connection_name As String = "connection name"
    Public Const closingTag As String = "</"
    Public Const connection As String = "connection"
    Public Const optionalExpression As String = "optionalExpression"
    Public Const states As String = "states"
    Public Const state As String = "state"
    Public Const stateId As String = "stateId"
    Public Const Threshold As String = "threshold"
    Public Const Data As String = "data"
    Public Const Animation As String = "animation"
    Public Const Animate As String = "animate"
    Public Const Color As String = "color"
    Public Const readFromTagExpressionRange As String = "readFromTagExpressionRange"
    Public Const constantExpressionRange As String = "constantExpressionRange"
    Public Const defaultExpressionRange As String = "defaultExpressionRange"
End Class

Public Enum EditCase
    Left
    Right
End Enum

Public Enum ECloseType
    Simple
    Complex
End Enum

Public Class ObjectTestClass
    Public Const caption As String = "caption"
    Public Const image As String = "image"
    Public Const connection As String = "connection"
    Public Const state As String = "state"
    Public Const threshold As String = "threshold"
    Public Const data As String = "data"
    Public Const animate As String = "animate"
    Public Const color As String = "color"
    Public Const readfromtagexpressionrange As String = "readFromTagExpressionRange"
    Public Const constantexpressionrange As String = "constantExpressionRange"
    Public Const defaultexpressionrange As String = "defaultExpressionRange"

End Class



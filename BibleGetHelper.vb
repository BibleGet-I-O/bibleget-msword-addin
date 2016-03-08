Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

Public Class BibleGetHelper

    Private indexes As Indexes = New Indexes()
    Public errorMessages As List(Of String) = New List(Of String)

    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Dim rgx As New Regex("''")
            myTranslation = rgx.Replace(myTranslation, "'")
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Public Function integrityCheck(ByVal myQuery As String, ByVal selectedVersions() As String) As Boolean
        Dim versionsStr As String = String.Join(",", selectedVersions)
        'Diagnostics.Debug.WriteLine("starting integrity check on myQuery <" & myQuery & ">")
        myQuery = New String(myQuery.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
        'Diagnostics.Debug.WriteLine("clean version of myQuery <" & myQuery & ">")
        'System.out.println("Starting integrity check on query "+myQuery+" for versions: "+versionsStr)

        '//final result is true until proved false
        '//set finFlag to false for non-breaking errors, or simply return false for breaking errors
        Dim finFlag As Boolean = True

        errorMessages.Clear()
        Dim queries As List(Of String) = New List(Of String)

        '//if english notation is found, translate to european notation
        If myQuery.Contains(":") And myQuery.Contains(".") Then
            errorMessages.Add(__("Mixed notations have been detected. Please use either english notation or european notation."))
            Return False
        ElseIf myQuery.Contains(":") Then
            If myQuery.Contains(",") Then
                myQuery = myQuery.Replace(",", ".")
            End If
            myQuery = myQuery.Replace(":", ",")
        End If

        If Not myQuery Is String.Empty Then
            If myQuery.Contains(";") Then
                '//System.out.println("We have a semicolon");
                queries.AddRange(myQuery.Split(";").ToList)
                queries.RemoveAll(Function(str) String.IsNullOrWhiteSpace(str))
            Else
                '//System.out.println("There is no semicolon");
                queries.Add(myQuery)
            End If
        End If

        Dim first As Boolean = True
        Dim currBook As String = ""

        If queries.Count = 0 Then
            errorMessages.Add(__("You cannot send an empty query."))
            Return False
        End If
        For Each querie As String In queries
            '//System.out.println(querie);
            querie = toProperCase(querie)
            '//System.out.println(querie);

            '//RULE 1: at least the first query must have a book indicator
            If first Then
                If Not Regex.IsMatch(querie, "^[1-3]{0,1}((\p{L}\p{M}*)+)(.*)") Then
                    errorMessages.Add(String.Format(__("The first query <{0}> in the querystring <{1}> must start with a valid book indicator!"), querie, myQuery))
                    finFlag = False
                End If
                first = False
            End If

            '//RULE 2: for every query that starts with a book indicator, 
            '//        the book indicator must be followed by valid chapter indicator;
            '//        else query must start with valid chapter indicator
            Dim bBooksContains As Integer
            Dim myidx As Integer = -1
            Dim tempBook As String = ""

            If Regex.IsMatch(querie, "^[1-3]{0,1}((\p{L}\p{M}*)+)(.*)") Then
                '//while we're at it, let's capture the book value from the query
                Dim pattern As String = "^[1-3]{0,1}((\p{L}\p{M}*)+)"
                'Matcher matcher = pattern.matcher(querie);
                Dim m As Match = Regex.Match(querie, pattern)
                If m.Success Then
                    tempBook = m.Groups(1).Value
                    'Diagnostics.Debug.WriteLine("Captured the book as <" & tempBook & ">")
                    bBooksContains = isValidBook(tempBook)
                    myidx = bBooksContains + 1
                    '//if(bBooksContains == -1 && bBooksAbbrevsContains == false){
                    If bBooksContains = -1 Then
                        errorMessages.Add(String.Format(__("The book indicator <{0}> in the query <{1}> is not valid. Please check the documentation for a list of valid book indicators."), tempBook, querie))
                        finFlag = False
                    Else
                        '//if(bBooksContains)
                        currBook = tempBook
                        '//querie = querie.replace(tempBook,"");
                    End If
                End If

                Dim pattern1 As String = "^[1-3]{0,1}((\p{L}\p{M}*)+)"
                Dim pattern2 As String = "^[1-3]{0,1}((\p{L}\p{M}*)+)[1-9][0-9]{0,2}"

                Dim count1 As Integer = 0
                For Each match1 As Match In Regex.Matches(querie, pattern1)
                    count1 += 1
                Next

                Dim count2 As Integer = 0
                For Each match2 As Match In Regex.Matches(querie, pattern2)
                    count2 += 1
                Next
                'Diagnostics.Debug.WriteLine("count1 = " & count1.ToString & " | count2 = " & count2.ToString)
                If Not Regex.IsMatch(querie, "^[1-3]{0,1}((\p{L}\p{M}*)+)[1-9][0-9]{0,2}(.*)") Or count1 <> count2 Then
                    errorMessages.Add(__("You must have a valid chapter following the book indicator!"))
                    finFlag = False
                End If
                querie = querie.Replace(tempBook, "")

            Else
                If Not Regex.IsMatch(querie, "^[1-9][0-9]{0,2}(.*)") Then
                    errorMessages.Add(__("A query that doesn't start with a book indicator must however start with a valid chapter indicator!"))
                    finFlag = False
                End If
            End If

            '//RULE 3: Queries with a dot operator must first have a comma operator; and cannot have more commas than dots
            If querie.Contains(".") Then
                Dim pattern11 As String = "[,|\-|\.][1-9][0-9]{0,2}\."
                If Not querie.Contains(",") Or Not Regex.IsMatch(querie, pattern11) Then
                    errorMessages.Add(__("You cannot use a dot without first using a comma or a dash. A dot is a liason between verses, which are separated from the chapter by a comma."))
                    finFlag = False
                End If

                Dim pattern3 As String = "(?<![0-9])(?=(([1-9][0-9]{0,2})\.([1-9][0-9]{0,2})))"
                'Matcher matcher3 = pattern3.matcher(querie);
                Dim count As Integer = 0
                For Each match3 As Match In Regex.Matches(querie, pattern3)
                    '//RULE 4: verse numbers around dot operators must be sequential
                    If CInt(match3.Groups(2).Value) >= CInt(match3.Groups(3).Value) Then
                        errorMessages.Add(String.Format(__("Verses concatenated by a dot must be consecutive, instead <{0}> is greater than or equal to <{1}> in the expression <{2}> in the query <{3}>"), match3.Groups(2).Value, match3.Groups(3).Value, match3.Groups(1).Value, querie))
                        finFlag = False
                    End If
                    count += 1
                Next
                '//RULE 5: Dot operators must be preceded and followed by a number from one to three digits, of which the first digit cannot be a 0                
                If count = 0 Or count <> querie.Count(Function(c As Char) c = ".") Then
                    errorMessages.Add(__("A dot must be preceded and followed by 1 to 3 digits of which the first digit cannot be zero.") + " <" + querie + ">")
                    finFlag = False
                End If
            End If

            '//RULE 6: Comma operators must be preceded and followed by a number from one to three digits, of which the first digit cannot be 0
            If querie.Contains(",") Then

                Dim pattern4 As String = "([1-9][0-9]{0,2})\,[1-9][0-9]{0,2}"
                'Matcher matcher4 = pattern4.matcher(querie);
                Dim count As Integer = 0
                Dim chapters As New List(Of Integer)
                For Each matcher4 As Match In Regex.Matches(querie, pattern4)
                    '//System.out.println("group0="+matcher4.group(0)+", group1="+matcher4.group(1));
                    chapters.Add(CInt(matcher4.Groups(1).Value))
                    count += 1
                Next
                If count = 0 Or count <> querie.Count(Function(c As Char) c = ",") Then
                    errorMessages.Add(__("A comma must be preceded and followed by 1 to 3 digits of which the first digit cannot be zero.") + " <" + querie + ">" + "(count=" + count.ToString + ",comma count=" + querie.Count(Function(c As Char) c = ",") + "); chapters=" + chapters.ToString())
                    finFlag = False

                Else
                    '// let's check the validity of the chapter numbers against the version indexes
                    '//for each chapter captured in the querystring
                    For Each chapter As Integer In chapters
                        If Not indexes.isValidChapter(chapter, myidx, selectedVersions.ToList) Then
                            Dim chapterLimit() As Integer = indexes.getChapterLimit(myidx, selectedVersions.ToList)
                            errorMessages.Add(String.Format(__("A chapter in the query is out of bounds: there is no chapter <{0}> in the book <{1}> in the requested version <{2}>, the last possible chapter is <{3}>"), chapter.ToString, currBook, String.Join(",", selectedVersions), String.Join(",", chapterLimit)))
                            finFlag = False
                        End If
                    Next
                End If
            End If

            If querie.Count(Function(c As Char) c = ",") > 1 Then
                If Not querie.Contains("-") Then
                    errorMessages.Add(__("You cannot have more than one comma and not have a dash!"))
                    finFlag = False
                End If
                Dim parts() As String = querie.Split("-")
                If parts.Length <> 2 Then
                    errorMessages.Add(__("You seem to have a malformed querystring, there should be only one dash."))
                    finFlag = False
                End If
                For Each p As String In parts
                    Dim pp(2) As Integer
                    Dim tt() As String = p.Split(",")
                    Dim x As Integer = 0
                    For Each t As String In tt
                        pp(x) = CInt(t)
                        x += 1
                    Next
                    If Not indexes.isValidChapter(pp(0), myidx, selectedVersions.ToList) Then
                        Dim chapterLimit() As Integer
                        chapterLimit = indexes.getChapterLimit(myidx, selectedVersions.ToList)
                        '//                        System.out.print("chapterLimit = ");
                        '//                        System.out.println(Arrays.toString(chapterLimit));
                        errorMessages.Add(String.Format(__("A chapter in the query is out of bounds: there is no chapter <{0}> in the book <{1}> in the requested version <{2}>, the last possible chapter is <{3}>"), pp(0).ToString, currBook, String.Join(",", selectedVersions), String.Join(",", chapterLimit)))
                        finFlag = False
                    Else
                        If Not indexes.isValidVerse(pp(1), pp(0), myidx, selectedVersions.ToList) Then
                            Dim verseLimit() As Integer = indexes.getVerseLimit(pp(0), myidx, selectedVersions.ToList)
                            '//                            System.out.print("verseLimit = ");
                            '//                            System.out.println(Arrays.toString(verseLimit));
                            errorMessages.Add(String.Format(__("A verse in the query is out of bounds: there is no verse <{0}> in the book <{1}> at chapter <{2}> in the requested version <{3}>, the last possible verse is <{4}>"), pp(1).ToString, currBook, pp(0).ToString, String.Join(",", selectedVersions), String.Join(",", verseLimit)))
                            finFlag = False
                        End If
                    End If
                Next


            ElseIf querie.Count(Function(c As Char) c = ",") = 1 Then
                Dim parts() As String = querie.Split(",")
                '//System.out.println(Arrays.toString(parts));
                If Not indexes.isValidChapter(CInt(parts(0)), myidx, selectedVersions.ToList) Then
                    Dim chapterLimit() As Integer = indexes.getChapterLimit(myidx, selectedVersions.ToList)
                    errorMessages.Add(String.Format(__("A chapter in the query is out of bounds: there is no chapter <{0}> in the book <{1}> in the requested version <{2}>, the last possible chapter is <{3}>"), parts(0), currBook, String.Join(",", selectedVersions), String.Join(",", chapterLimit)))
                    finFlag = False

                Else
                    Dim highverse As Integer
                    If parts(1).Contains("-") Then
                        Dim highverses As New Stack(Of Integer)
                        Dim pattern11 As String = "[,\.][1-9][0-9]{0,2}\-([1-9][0-9]{0,2})"
                        'Matcher matcher11 = pattern11.matcher(querie);
                        For Each matcher11 As Match In Regex.Matches(querie, pattern11)
                            highverses.Push(CInt(matcher11.Groups(1).Value))
                        Next
                        If highverses.Count Then
                            highverse = highverses.Pop()
                            If Not indexes.isValidVerse(highverse, CInt(parts(0)), myidx, selectedVersions.ToList) Then
                                Dim verseLimit() As Integer = indexes.getVerseLimit(CInt(parts(0)), myidx, selectedVersions.ToList)
                                errorMessages.Add(String.Format(__("A verse in the query is out of bounds: there is no verse <{0}> in the book <{1}> at chapter <{2}> in the requested version <{3}>, the last possible verse is <{4}>"), highverse, currBook, parts(0), String.Join(",", selectedVersions), String.Join(",", verseLimit)))
                                finFlag = False
                            End If
                        Else
                            highverse = Nothing
                            Dim verseLimit() As Integer = indexes.getVerseLimit(CInt(parts(0)), myidx, selectedVersions.ToList)
                            errorMessages.Add(String.Format(__("A verse in the query is out of bounds: there is no verse <{0}> in the book <{1}> at chapter <{2}> in the requested version <{3}>, the last possible verse is <{4}>"), highverse, currBook, parts(0), String.Join(",", selectedVersions), String.Join(",", verseLimit)))
                            finFlag = False
                        End If
                    Else
                        Dim pattern12 As String = ",([1-9][0-9]{0,2})"
                        'Matcher matcher12 = pattern12.matcher(querie);
                        highverse = -1
                        For Each match As Match In Regex.Matches(querie, pattern12)
                            highverse = CInt(match.Groups(1).Value)
                            '//System.out.println("[line 376]:highverse="+Integer.toString(highverse));
                        Next
                        If highverse <> -1 Then
                            '//System.out.println("Checking verse validity for book "+myidx+" chapter "+parts[0]+"...");
                            If Not indexes.isValidVerse(highverse, CInt(parts(0)), myidx, selectedVersions.ToList) Then
                                Dim verseLimit() As Integer = indexes.getVerseLimit(CInt(parts(0)), myidx, selectedVersions.ToList)
                                errorMessages.Add(String.Format(__("A verse in the query is out of bounds: there is no verse <{0}> in the book <{1}> at chapter <{2}> in the requested version <{3}>, the last possible verse is <{4}>"), highverse.ToString, currBook, parts(0), String.Join(",", selectedVersions), String.Join(",", verseLimit)))
                                finFlag = False
                            End If
                        End If
                    End If

                    Dim pattern13 As String = "\.([1-9][0-9]{0,2})$"
                    'Matcher matcher13 = pattern13.matcher(querie);
                    highverse = -1
                    For Each match As Match In Regex.Matches(querie, pattern13)
                        highverse = CInt(match.Groups(1).Value)
                    Next
                    If highverse <> -1 Then
                        If Not indexes.isValidVerse(highverse, CInt(parts(0)), myidx, selectedVersions.ToList) Then
                            Dim verseLimit() As Integer = indexes.getVerseLimit(CInt(parts(0)), myidx, selectedVersions.ToList)
                            errorMessages.Add(String.Format(__("A verse in the query is out of bounds: there is no verse <{0}> in the book <{1}> at chapter <{2}> in the requested version <{3}>, the last possible verse is <{4}>"), highverse, currBook, parts(0), String.Join(",", selectedVersions), String.Join(",", verseLimit)))
                            finFlag = False
                        End If
                    End If
                End If


            Else  '//if there's no comma, it's either a single chapter or an extension of chapters with a dash
                '//System.out.println("no comma found");
                Dim parts() As String = querie.Split("-")
                '//System.out.println(Arrays.toString(parts));
                Dim highchapter As Integer = CInt(parts(parts.Length - 1))
                If Not indexes.isValidChapter(highchapter, myidx, selectedVersions.ToList) Then
                    Dim chapterLimit() As Integer = indexes.getChapterLimit(myidx, selectedVersions.ToList)
                    errorMessages.Add(String.Format(__("A chapter in the query is out of bounds: there is no chapter <{0}> in the book <{1}> in the requested version <{2}>, the last possible chapter is <{3}>"), highchapter.ToString, currBook, String.Join(",", selectedVersions), String.Join(",", chapterLimit)))
                    finFlag = False
                End If
            End If

            If querie.Contains("-") Then
                '//RULE 7: If there are multiple dashes in a query, there cannot be more dashes than there are dots minus 1
                Dim dashcount As Integer = querie.Count(Function(c As Char) c = "-")
                Dim dotcount As Integer = querie.Count(Function(c As Char) c = ".")
                If dashcount > 1 Then
                    If dashcount - 1 > dotcount Then
                        errorMessages.Add(__("There are multiple dashes in the query, but there are not enough dots. There can only be one more dash than dots.") + " <" + querie + ">")
                        finFlag = False
                    End If
                End If

                '//RULE 8: Dash operators must be preceded and followed by a number from one to three digits, of which the first digit cannot be 0
                Dim pattern5 As String = "([1-9][0-9]{0,2}\-[1-9][0-9]{0,2})"
                'Matcher matcher5 = pattern5.matcher(querie);
                Dim count As Integer = 0
                For Each match As Match In Regex.Matches(querie, pattern5)
                    count += 1
                Next
                If count = 0 Or count <> querie.Count(Function(c As Char) c = "-") Then
                    errorMessages.Add(__("A dash must be preceded and followed by 1 to 3 digits of which the first digit cannot be zero.") + " <" + querie + ">")
                    finFlag = False
                End If

                '//RULE 9: If a comma construct follows a dash, there must also be a comma construct preceding the dash
                Dim pattern6 As String = "\-([1-9][0-9]{0,2})\,"
                'Matcher matcher6 = pattern6.matcher(querie);
                If Regex.IsMatch(querie, pattern6) Then
                    Dim pattern7 As String = "\,[1-9][0-9]{0,2}\-"
                    'Matcher matcher7 = pattern7.matcher(querie);
                    If Not Regex.IsMatch(querie, pattern7) Then
                        errorMessages.Add(__("If there is a chapter-verse construct following a dash, there must also be a chapter-verse construct preceding the same dash.") + " <" + querie + ">")
                        finFlag = False

                    Else
                        '//RULE 10: Chapters before and after dashes must be sequential
                        Dim chap1 As Integer = -1
                        Dim chap2 As Integer = -1

                        Dim pattern8 As String = "([1-9][0-9]{0,2})\,[1-9][0-9]{0,2}\-"
                        'Matcher matcher8 = pattern8.matcher(querie);
                        Dim match8 As Match = Regex.Match(querie, pattern8)
                        If match8.Success Then
                            chap1 = CInt(match8.Groups(1).Value)
                        End If
                        Dim pattern9 As String = "\-([1-9][0-9]{0,2})\,"
                        'Matcher matcher9 = pattern9.matcher(querie);
                        Dim match9 As Match = Regex.Match(querie, pattern9)
                        If match9.Success Then
                            chap2 = CInt(match9.Groups(1).Value)
                        End If

                        If chap1 >= chap2 Then
                            errorMessages.Add(String.Format(__("Chapters must be consecutive. Instead the first chapter indicator <{0}> is greater than or equal to the second chapter indicator <{1}> in the expression <{2}>"), chap1.ToString, chap2.ToString, querie))
                            finFlag = False
                        End If
                    End If

                Else

                    '//if there are no comma constructs immediately following the dash
                    '//RULE 11: Verses (or chapters if applicable) around each of the dash operator(s) must be sequential
                    Dim pattern10 As String = "([1-9][0-9]{0,2})\-([1-9][0-9]{0,2})"
                    'Matcher matcher10 = pattern10.matcher(querie);
                    For Each match As Match In Regex.Matches(querie, pattern10)
                        Dim num1 As Integer = CInt(match.Groups(1).Value)
                        Dim num2 As Integer = CInt(match.Groups(2).Value)
                        If num1 >= num2 Then
                            errorMessages.Add(String.Format(__("Verses (or chapters if applicable) around the dash operator must be consecutive. Instead <{0}> is greater than or equal to <{1}> in the expression <{2}>"), num1.ToString, num2.ToString, querie))
                            finFlag = False
                        End If
                    Next

                End If
            End If

        Next


        Return finFlag

    End Function


    Public Function toProperCase(ByVal txt As String) As String
        Dim idx As Integer = 0
        While Not Regex.IsMatch(Char.ToString(txt.Chars(idx)), "[a-zA-Z]")
            If idx = txt.Length - 1 Then Exit While
            idx += 1
        End While
        If idx < txt.Length - 2 Then
            Return txt.Substring(0, idx) + Char.ToString(txt.Chars(idx)).ToUpper + txt.Substring(idx + 1).ToLower
        Else
            Return txt
        End If
    End Function

    Public Function isValidBook(ByVal book As String) As Integer
        Try
            Dim bibleGetDB As New BibleGetDatabase
            'Dim biblebooks As List(Of String(,))
            Dim biblebooks As New JArray
            For i As Integer = 0 To 72
                Dim usrprop As String = bibleGetDB.getMetaData("BIBLEBOOKS" + i.ToString)
                '//System.out.println("value of BIBLEBOOKS"+Integer.toString(i)+": "+usrprop);                
                'JsonReader jsonReader = Json.createReader(new StringReader(usrprop));
                Dim jRRArray As JArray = JArray.Parse(usrprop)
                'Dim jRRArray As JArray = JArray.FromObject(jsObj)
                biblebooks.Add(jRRArray)
                'biblebooks.Add(jsObj.Values(Of String(,)).ToArray)
            Next

            'JsonArray biblebooks = biblebooksBldr.build();
            If biblebooks.Count > 0 Then
                Return idxOf(book, biblebooks)
            End If
        Catch ex As Exception
            'Logger.getLogger(HTTPCaller.class.getName()).log(Level.SEVERE, null, ex);
            Diagnostics.Debug.WriteLine(ex.Message)
        End Try
        Return -1
    End Function

    Public Function idxOf(ByVal needle As String, ByVal haystack As JArray) As Integer
        Dim count As Integer = 0
        For Each m As JArray In haystack
            'Dim m As JArray = i
            If m(0).GetType().IsArray Then
                For Each x As JValue In m
                    '//System.out.println("looking for '"+needle+"' in "+x.toString());
                    If x.ToString().Contains("""" + needle + """") Then Return count
                Next
            Else
                If m.ToString().Contains("""" + needle + """") Then Return count
            End If
            count += 1
        Next
        Return -1
    End Function

End Class

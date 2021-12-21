# MMRTokenizerXL
## 1.Tokenizer for Myanmar Unicode Pyidaungsu Font (Visual Order)
![MMRTokenizerXL](/images/MMRTokenizerXL.png)
## 2.The Journey
As usual, while I was haunting in my usual Facebook Excel groups, I came across somebody asking a question on how to sort a Myanmar Name in Myanmar font using the last name.
As we, all Myanmar people, know, our names do not have a first name/last name basis but rather like a jumble of nice words joined together to get a beautiful or following the Myanmar astrological beliefs.\
The main problem with this naming system concerns something else.

It is the way we write our language.\
We don't really need to use a white space to separate each word or ?syllable. I am not a linguist. So, I don't know how it is called.\
Anyway, the real problem is that I found it hard to find out the word break points. I mean the exact point where one word of a name ends and where the next word starts.

Let's consider a name in Myanmar/Burmese e.g. ခင်မို့မို့အောင် (=Khin Moh Moh Aung) this is a female's name.\
In this case, even though we spell in English like so inside the parentheses, the Myanmar words do NOT have any white space between them. They just don't have to.\
If we want to separate/tokenize/split the name in English, we could easily do so in VBA using the SPLIT function or use some formulas in Excel worksheet UI.\
But we have no way of achieving the same goal with a Myanmar font, the Pyidaungsu font, namely, because it is not easy to programmatically identify where the consonants are, using VBA.

I am so interested in Natural Language Process (NLP) but it is very hard for me to understand those papers written by NLP specialists because I don't have the basics of linguistics and I was never officially trained in programming languages, especially the VBA, I have currently and mostly written in.\
However, I persevered and read through many papers but I got no where.\
Then I found some github pages where some people wrote up something about a python package and they showed an image which shows, in turn, a tokenized Myanmar words and I was so addicted to it.

Now that there is this question about how to sort Myanmar names using the last word of the name, I realized that it would be best to get the last consonant out and sort by that, much like we do with English names.\
So I tried to get the consonant of the last word from a name. It was very hard for me because I didn't know how Pyidaungsu font stitches each parts of a word.

For example, for the same name, ခင်မို့မို့အောင်, (a beautiful name for a girl btw), we actually typed it up using the Windows10's Burmese Keyboard (Visual Order) in Pyidaungsu font like (ခ, င, ် , မ, ိ, ု, ့ , မ, ိ, ု, ့ , ‌ေ, အ, ာ, င, ်     ) because we need to type the way we spell as well as the way we see/read it. There is another Burmese Keyboard using Phonetic Order but I am NOT using it and I don't develop code for that because the typing is pretty hard on it.\
From trial and error and trying to convert each typed parts into Unicode values using AscW and ChrW, I realized that, the developers of the Pyidaungsu font decided to put together a system (albeit an incomplete one, more on this later) which, in a way, places the consonants in front of every other appendages/diacritics? of a Burmese word spelling.

What I meant is for the last part of the name, အောင်, they placed the consonant အ in front of the other stuff. And they bind them together like အ, ‌ေ, ာ, င, ် , and if that is a uniform method, I could come up with a UDF for tokenization quite easily!
So I tried to loop from the right-most part of the name and bam, the first obstacle hits me, right in the face.\
The င under the ် , the င is actually a consonant but it must be combined with the Athat/diacritic? or ် , to create the whole word.
So I had to ignore the first consonant that comes after the Athat.\
The real problem now begins because Myanmar words have other part that comes between the Athat and the consonant!!! And the number of stuff that can come between is variable and can be more than 2 or 3.

I wrote incompleteness above because in some cases like, သန့်, we typed it up like သ, န, ် , ့   so I was thinking like the Athat ်   would come second last and the အောက်မြစ် ့  would come after it and I was totally wrong.
In reality, the Athat ်  comes rightmost then the ့  comes after that which wreaks havoc in my algorithm and I had to spend hours trying to fix that.\
Therefore, I had to come up with a way to filter through, to reach the consonant of the last word in a name. And I did.\
The algorithm is not graceful nor very cool but it works!

I was like OK, I can get the consonant of the last word in a name already. Why not get all the consonants using this algorighm.\
By this time, I can sort a column using the consonant of the last word of a name!\
I expanded and improved my code to be able extract all the consonants of a name!!!

The idea now was that all the components of each words of a name must be between the consonants! because Pyidaungsu font was made to combine words like that! Thank you, Pyidaungsu font dev team!!\
So, I just had to put a separator/delimited just before each consonant and then BAM! the name is tokenized!

Then I found that using the consonant of the last word of a name for filtering is not very accurate because in Burmese language, we have a system of laying out the alphabetical order of words, which we called as, မြန်မာအက္ခရာစဉ်. This topic was taungt in the first year of my bachelor's degree and I failed this subject, the Myanmar language!\
I tried using the consonant extraction UDF and compared it to sorting by using only the consonant from the last word. And I found that the difference is quite impactful.\
This drives me to finish the whole Tokenization UDF because the Burmese language sorting system uses not just consonants but other parts in the building of a word like those I already mentioned above for example, ်  or ာ or ိ or ု or ့  etc of may of those things/diacritics?.\
Enough said and I wish to admit that I used just string manipulation functions in this UDF rather than the NLP methods. I am not employing ML or AI with all this.

## 3.The UDFs
The source code of the UDFs may be released as plain text.\
There are currently 4 UDFs in the .bas and .xlsm files upon release.
1. [MMRTokenizer](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#11mmrtokenizer)
2. [MMRManipulator](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#12mmrmanipulator)
3. [getMMRConsonants](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#13getmmrconsonants)
4. [MMRParser](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#14mmrparser)

### 3.1.MMRTokenizer
MMRTokenizer is designed to be used mainly for tokenization of Myanmar words without additional bells and whistles, as this UDF was purported to be used for further processing into NLP methods, rather than intended for general everyday use.\
If extra functionality is required, the users are encouraged to use [MMRManipulator UDF](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#mmrmanipulator).\
Since this UDF is mainly intended for NLP-related usage, it's users are expected to be able to manipulate the VBA source code directly to change the separator to their whims, so no switching arguments are included for that purpose.

Users can directly copy the UDF code below instead of downloading the .xlsm or .bas modules from [Releases Section](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases).
```VBA
Option Explicit
'**********************************************************************************************************************************
'*Users of the following VBA code are not allowed to share the code commercially without written approval from the developer.     *
'*Any commercial distribution of the code herein requires acknowledgement, consent and approval from the author.                  *
'*The developer of the code holds complete and thorough copyrights, however, no authorization is required for educational and     *
'*humanitarian uses, in which case, this whole declaration section must be included wheresoever the code herein is placed.        *
'*Failure to comply with above declarations shall be liable to the full extent of the law.                                        *
'*The VBA code provided herewith has no guarantee whatsoever with it and any untoward effect(s) that occur(s) shall not be held   *
'*liable to the developer and it is taken as a legally binding fact that the user(s) of said code must have agreed to this        *
'*disclaimer, in order to use it.                                                                                                 *
'*Contact info can be found at https://github.com/4R3B3LatH34R7                                                                   *
'**********************************************************************************************************************************

'Can place the constants in each function if only some functions were required
Public Const kagyi = 4096
Public Const ah = 4129 '+9 to include ou
Public Const athat = 4154
Public Const shiftF = 4153 'for typing something under something
Public Const witecha = 4140
Public Const moutcha = 4139

'Return a tokenized Myanmar String
Function MMRTokenizer(target As Range) As String
Dim ch As String
Dim returnString As String
Dim charCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
Dim previousCharAt As Long
    If target.Cells.CountLarge > 1 Then MMRTokenizer = ">1Cell!": Exit Function
    returnString = "": previousChIsAthat = False: shiftFfound = False: previousCharAt = Len(target.Value) + 1
    If target.CountLarge = 1 Then
        If target.Value <> "" Then
            For charCounter = Len(target.Value) To 1 Step -1
                ch = Mid(target.Value, charCounter, 1)
                If AscW(ch) <> shiftF Then
                    If Not shiftFfound Or AscW(ch) = athat Then
                        If AscW(ch) <> athat Then
                            If AscW(ch) >= kagyi And AscW(ch) < ah + 9 Then
                                If Not previousChIsAthat Then
                                    returnString = Mid(target.Value, charCounter, previousCharAt - charCounter) & IIf(Len(returnString) > 0, "|", "") & returnString
                                    previousCharAt = charCounter
                                Else
                                    previousChIsAthat = False
                                End If
                            Else
                                If AscW(ch) = witecha Or AscW(ch) = moutcha Then
                                    previousChIsAthat = False
                                End If
                            End If
                        Else
                            previousChIsAthat = True
                            If shiftFfound Then shiftFfound = False
                        End If
                    Else
                        shiftFfound = False
                        If previousChIsAthat Then previousChIsAthat = False
                    End If
                Else
                    shiftFfound = True
                End If
            Next charCounter
        End If
    End If
    MMRTokenizer = returnString
End Function
```

### 3.2.MMRManipulator
This is a tool spawned from being able to tokenize Myanmar/Burmese words typed using Burmese Visual Order Keyboard in Windows, deviating from this would result in lesser performance.\
It can be used to tokenize the words, for any purpose, like for sorting, counting, replacing...etc...with the sky at the limit of the users' imagination.\
With argument switch(es), users can change the tokenization character to become anything, any text, any string, even nothing!\
The users can also reverse the whole Myanmar sentence word by word with the first word becoming the last word and vice versa...\
The result of using this tool can be found in the [photo](/images/MMRTokenizerXL.png) under the right most column (Column F).
If cell A1 contains မိုးအောင်ခင်, then calling from inside cell B1 like, =MMRTokenizer(A1), shall return မိုး|အောင်|ခင်.

The following steps can guide users use MMRManipulator to sort Myanmar names, words, sentences in reverse in the following short .gif.
![MMRManipulator](/images/MMRManipulator_sorting.gif)
The process is simple in that, users just need to use the UDF to reverse the range containing Myanmar words.\
The UDF only requires the first argument, out of the existing 5: 
1. the target range, containing the target text string, is essential
2. the second argument is for defining the left or starting wrapper and can be anything text/string, for example, "(" or "<" or "{" or "\[".
3. the third argument is for defining the separator/delimiter for the output of the UDF, which can be anything from ""(vbNullString in VBA) or blank cell (in Excel) or a space or any word in Burmese or English and since no check was performed on this argument's validity, it can be quite powerful and dangerous at the same time.
4. the fourth argument is for defining the right or ending wrapper and can be anything text/string, including but not limited to e.g., ")" or ">" or "}" or "]".
5. the fifth argument is a boolean variable which acts as a switch for reversing the output of the UDF.

So, if cell A1 contains "ကိုကိုအေး" and from inside cell B1, if we call the UDF as: (let "->" denotes "returns")
1. =MMRManipulator(A1) -> ကို|ကို|အေး
2. =MMRManipulator(A1,,"@",,) -> ကို@ကို@အေး
3. =MMRManipulator(A1,,"",,TRUE) -> အေးကိုကို (please note that "" is not space but denotes nothing)

Apart from the target cell reference (the 1st argument), the remaining 4 arguments are optional, thus, calling like =MMRManipulator(A1) or =MMRManipulator(A1,,,,) is legitimate and will return ကို|ကို|အေး anyway.\
The reason behind including starting and ending wrappers is to make the output more in line with lists in other programming languages like Python.\
For example, users can call the UDF as =MMRManipulator(A1,"""",",",CHAR(34)) will return "ကို","ကို","အေး" and please note here that the Double quote character must be called as 4xDouble Quotes or CHAR(34) which is a requirement of VBA. Another good example would be calling like =MMRManipulator(A1,"(","-",")",TRUE) which would return something like (အေး)-(ကို)-(ကို).

Users can directly copy the UDF code below instead of downloading the .xlsm or .bas modules from [Releases Section](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases).
```VBA
Option Explicit
'**********************************************************************************************************************************
'*Users of the following VBA code are not allowed to share the code commercially without written approval from the developer.     *
'*Any commercial distribution of the code herein requires acknowledgement, consent and approval from the author.                  *
'*The developer of the code holds complete and thorough copyrights, however, no authorization is required for educational and     *
'*humanitarian uses, in which case, this whole declaration section must be included wheresoever the code herein is placed.        *
'*Failure to comply with above declarations shall be liable to the full extent of the law.                                        *
'*The VBA code provided herewith has no guarantee whatsoever with it and any untoward effect(s) that occur(s) shall not be held   *
'*liable to the developer and it is taken as a legally binding fact that the user(s) of said code must have agreed to this        *
'*disclaimer, in order to use it.                                                                                                 *
'*Contact info can be found at https://github.com/4R3B3LatH34R7                                                                   *
'**********************************************************************************************************************************
'Can place the constants in each function if only some functions were required
Public Const kagyi = 4096
Public Const ah = 4129 '+9 to include ou
Public Const athat = 4154
Public Const shiftF = 4153 'for typing something under something
Public Const witecha = 4140
Public Const moutcha = 4139

'Return tokenized words using user-selectable optional separator and ability to reverse the Myanmar word string
Function MMRManipulator( _
                        target As Range, _
                        Optional lWrapper As String = "", _
                        Optional separator As String = "|", _
                        Optional rWrapper As String = "", _
                        Optional reversed As Boolean = False) As String
Dim ch As String
Dim returnString As String
Dim charCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
Dim previousCharAt As Integer '?long
Const defaultSeparator As String = "|"
    If target.Cells.CountLarge > 1 Then MMRManipulator = ">1Cell!": Exit Function
    returnString = "": previousChIsAthat = False: shiftFfound = False: previousCharAt = Len(target.Value) + 1
    If target.CountLarge = 1 Then
        If target.Value <> "" Then
            For charCounter = Len(target.Value) To 1 Step -1
                ch = Mid(target.Value, charCounter, 1)
                If AscW(ch) <> shiftF Then
                    If Not shiftFfound Or AscW(ch) = athat Then
                        If AscW(ch) <> athat Then
                            If AscW(ch) >= kagyi And AscW(ch) < ah + 9 Then
                                If Not previousChIsAthat Then
                                    returnString = IIf(reversed, returnString, Mid(target.Value, charCounter, previousCharAt - charCounter)) & _
                                                   IIf(Len(returnString) > 0, defaultSeparator, "") & _
                                                   IIf(reversed, Mid(target.Value, charCounter, previousCharAt - charCounter), returnString)
                                    previousCharAt = charCounter
                                Else
                                    previousChIsAthat = False
                                End If
                            Else
                                If AscW(ch) = witecha Or AscW(ch) = moutcha Then
                                    previousChIsAthat = False
                                End If
                            End If
                        Else
                            previousChIsAthat = True
                            If shiftFfound Then shiftFfound = False
                        End If
                    Else
                        shiftFfound = False
                        If previousChIsAthat Then previousChIsAthat = False
                    End If
                Else
                    shiftFfound = True
                End If
            Next charCounter

            If InStr(returnString, defaultSeparator) > 0 Then 'check for names like may??
                returnString = Replace(returnString, defaultSeparator, separator)
            End If
            returnString = lWrapper & Join(Split(returnString, separator), rWrapper & separator & lWrapper) & rWrapper

        End If
    End If
    MMRManipulator = returnString
End Function
```

### 3.3.getMMRConsonants
This UDF was designed in the earlier stages of development of MMRTokenizer to help me identify, check and confirm the location of Myanmar consonants in a cell containing Myanmar word(s).\
There are altogether 4 possible arguments that can be passed when calling it.
1. target range (required)
2. reversed order (optional with default=false)
3. last character only (optional with default=false)
4. location of consonants (optional with default=false)
Apart from the target range, the rest are optional.\
The arguments are pretty obvious and I believe that there is no need for further explanation.\
The output of this UDF can be seen in the [photo](/images/MMRTokenizerXL.png) under the Column C.

Users can directly copy the UDF code below instead of downloading the .xlsm or .bas modules from [Releases Section](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases).
```VBA
Option Explicit
'**********************************************************************************************************************************
'*Users of the following VBA code are not allowed to share the code commercially without written approval from the developer.     *
'*Any commercial distribution of the code herein requires acknowledgement, consent and approval from the author.                  *
'*The developer of the code holds complete and thorough copyrights, however, no authorization is required for educational and     *
'*humanitarian uses, in which case, this whole declaration section must be included wheresoever the code herein is placed.        *
'*Failure to comply with above declarations shall be liable to the full extent of the law.                                        *
'*The VBA code provided herewith has no guarantee whatsoever with it and any untoward effect(s) that occur(s) shall not be held   *
'*liable to the developer and it is taken as a legally binding fact that the user(s) of said code must have agreed to this        *
'*disclaimer, in order to use it.                                                                                                 *
'*Contact info can be found at https://github.com/4R3B3LatH34R7                                                                   *
'**********************************************************************************************************************************
'Can place the constants in each function if only some functions were required
Public Const kagyi = 4096
Public Const ah = 4129 '+9 to include ou
Public Const athat = 4154
Public Const shiftF = 4153 'for typing something under something
Public Const witecha = 4140
Public Const moutcha = 4139

'Return all consonants within a range with optional reversing, last character only or consonant locations instead of actual ones
Function getMMRConsonants(target As Range, Optional reversedOrder As Boolean = False, Optional lastCharOnly As Boolean = False, Optional LOC As Boolean = False) As String
Dim ch As String
Dim returnString As String
Dim charCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
    If target.Cells.CountLarge > 1 Then getMMRConsonants = ">1Cell!": Exit Function
    returnString = "": previousChIsAthat = False: shiftFfound = False
    If target.CountLarge = 1 Then
        If target.Value <> "" Then
            For charCounter = Len(target.Value) To 1 Step -1
                ch = Mid(target.Value, charCounter, 1)
                If AscW(ch) <> shiftF Then
                    If Not shiftFfound Or AscW(ch) = athat Then
                        If AscW(ch) <> athat Then
                            If AscW(ch) >= kagyi And AscW(ch) < ah + 9 Then
                                If Not previousChIsAthat Then
                                    returnString = IIf(LOC, CStr(charCounter) + IIf(Len(returnString) > 0, "|", "") + returnString, ch + returnString)
                                Else
                                    previousChIsAthat = False
                                End If
                            Else
                                If AscW(ch) = witecha Or AscW(ch) = moutcha Then
                                    previousChIsAthat = False
                                End If
                            End If
                        Else
                            previousChIsAthat = True
                            If shiftFfound Then shiftFfound = False
                        End If
                    Else
                        shiftFfound = False
                        If previousChIsAthat Then previousChIsAthat = False
                    End If
                Else
                    shiftFfound = True
                End If
            Next charCounter
        End If
    End If
    getMMRConsonants = IIf(lastCharOnly, Right(returnString, 1), IIf(reversedOrder, StrReverse(returnString), returnString))
End Function
```

### 3.4.MMRParser
This UDF was also written in the earlier part of the development of MMRTokenizer to help me confirm the location of the Myanmar consonants.\
Only 1 argument is required, out of the 3 possible arguments.\
1. target range (required)
2. output as Myanmar (optional with default=false)
3. hightlight Consonants (optional with default=false)

This UDF just returns the Unicode values (numbers) as a string of text. For example, if cell A1 contains သီရိ, calling from inside cell B1 as,
1. =MMRParser(A1) -> 4126|4142|4123|4141 -> returning the Unicode values separated by | (=pipe) character.
2. =MMRParser(A1,TRUE) -> သ|ီ|ရ|ိ -> showing how the word was spelled.
3. =MMRParser(A1,,TRUE) -> returns သ|4142|ရ|4141 -> showing the location of Myanmar consonants in the word.

I don't think there would be much use for this UDF by everyday users, however, I am hoping that it would be useful to NLP devs.
The output of this UDF can be seen in the [photo](/images/MMRTokenizerXL.png) under the Column D.

Users can directly copy the UDF code below instead of downloading the .xlsm or .bas modules from [Releases Section](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases).
```VBA
Option Explicit
'**********************************************************************************************************************************
'*Users of the following VBA code are not allowed to share the code commercially without written approval from the developer.     *
'*Any commercial distribution of the code herein requires acknowledgement, consent and approval from the author.                  *
'*The developer of the code holds complete and thorough copyrights, however, no authorization is required for educational and     *
'*humanitarian uses, in which case, this whole declaration section must be included wheresoever the code herein is placed.        *
'*Failure to comply with above declarations shall be liable to the full extent of the law.                                        *
'*The VBA code provided herewith has no guarantee whatsoever with it and any untoward effect(s) that occur(s) shall not be held   *
'*liable to the developer and it is taken as a legally binding fact that the user(s) of said code must have agreed to this        *
'*disclaimer, in order to use it.                                                                                                 *
'*Contact info can be found at https://github.com/4R3B3LatH34R7                                                                   *
'**********************************************************************************************************************************
'Can place the constants in each function if only some functions were required
Public Const kagyi = 4096
Public Const ah = 4129 '+9 to include ou
Public Const athat = 4154
Public Const shiftF = 4153 'for typing something under something
Public Const witecha = 4140
Public Const moutcha = 4139

'Parse Myanmar strings into Unicode code values or Myanmar consonants and Diacritics and can return consonants in Burmese combined with numerical Diacritics
Function MMRParser(target As Range, Optional outputMMR As Boolean = False, Optional highlightConsonants As Boolean = False) As String
Dim returnStringArray()
Dim ch As String
Dim chCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
Dim legitConsonantFound As Boolean
    If target.Cells.CountLarge > 1 Then MMRParser = ">1Cell!": Exit Function
    previousChIsAthat = False: shiftFfound = False
    If target.CountLarge = 1 Then
        If target.Value <> "" Then
            ReDim returnStringArray(1 To Len(target.Value))
            For chCounter = Len(target.Value) To 1 Step -1
                ch = Mid(target.Value, chCounter, 1)
                legitConsonantFound = False
                If AscW(ch) <> shiftF Then
                    If Not shiftFfound Or AscW(ch) = athat Then
                        If AscW(ch) <> athat Then
                            If AscW(ch) >= kagyi And AscW(ch) < ah + 9 Then
                                If Not previousChIsAthat Then
                                    returnStringArray(chCounter) = IIf(outputMMR, ch, IIf(highlightConsonants, ch, AscW(ch)))
                                    legitConsonantFound = True
                                Else
                                    previousChIsAthat = False
                                End If
                            Else
                                If AscW(ch) = witecha Or AscW(ch) = moutcha Then
                                    previousChIsAthat = False
                                End If
                            End If
                        Else
                            previousChIsAthat = True
                            If shiftFfound Then shiftFfound = False
                        End If
                    Else
                        shiftFfound = False
                        If previousChIsAthat Then previousChIsAthat = False
                    End If
                Else
                    shiftFfound = True
                End If
                If Not legitConsonantFound Then returnStringArray(chCounter) = IIf(outputMMR, ch, AscW(ch))
            Next chCounter
        End If
    End If
    MMRParser = Join(returnStringArray, "|")
End Function
```
## 3.5.MMRTokenizer v1.2
MMRTokenizer v1.2 was released on 20DEC2021 and can be found in the [Releases Section](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases/tag/v1.2\).
It contains several new functions and some optimizations.\
New functions were added to aid in the general usability of Myanmar users who might find difficulties with:
1. Gender Identification/Classification/Analyses
2. Getting a count of words in a Myanmar Text String
3. Further manipulation of Myanmar Text Strings apart from MMRManipulator UDF from the previous version

The following photo is just an introduction to how the new functions could be used.
![v1.2New Functions](/images/mmrtokenizer_newFuncs_v1.2.png)
The green conditional highlight was used to show that correct number of columns/cells were selected with CSE before calling MMRSplit UDF.\
The orange highlighted cells containing #N/A should be easily noticeable as cells which are part of an array output but there was no values in the array for these areas, meaning extra columns/cells were selected while entering the UDF with CSE.

### 3.5.1.New Functions
There are altogether 5 new functions in v1.2. <b>All of these functions require MMRManipulator UDF.</b>\
1. [MMRSplit](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#1mmrsplittarget-as-rangeas-variant-string)
2. [MMRLen](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#2mmrlentarget-as-rangeas-long)
3. [MMRLeft](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#3mmrlefttarget-as-range-howmany-as-longas-string)
4. [MMRRight](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#4mmrrighttarget-as-range-howmany-as-longas-string)
5. [MMRMid](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#5mmrmidtarget-as-range-startpos-as-long-howmany-as-longas-string)

#### 3.5.1.1.MMRSplit(target as Range)as Variant 'String
This UDF is actually best used with Office365-Excel on a Windows computer.\
The reason behind this is, that, it splits a Myanmar word like a name or a sentence into it's component words (<b>NOT</b>consonants/diacritics etc) into ajacent cells (because it is an array formula). This feature is best suited to be used in a Excel365 environment on a Windows computer.\
In earlier versions of Excel, a CSE is required to enter this formula as an array formula.\
If no such precedence were performed, there will be N/A errors which could be avoided by using the next new formula, MMRLen.
````VBA
'Requires MMRManipulator
'Splits a Myanmar Text String into its component words and returns them as an array into adjacent cells, must use CSE except on Office365 Windows
Function MMRSplit(target As Range) As Variant
    If target.Cells.CountLarge > 1 Then MMRSplit = ">1Cell!": Exit Function
    MMRSplit = Split(MMRManipulator(target), "|")
End Function
````

#### 3.5.1.2.MMRLen(target as Range)as Long
The new functions in v1.2 are created to mimic the default string functions in Excel and VBA like Split (VBA only), Len, Left, Right, Mid etc. of string manipulation functions.\
The MMRLen function would simple return the length of a Text String in Myanmar Language typed using Pyidaungsu Font with Burmese Visual Order keyboard.\
Be mindful that he return from MMRLen is not going to be the same as the Len function/formula.\
If Cell A1 contains ABC then =Len(A1) would produce 3.\
Contrary, if Cell A2 contains "အောင်မြင့်", then =Len(A1) would produce 10 because it was spelled like ‌ေ,အ,ာ,င,်, မ,ြ,င,့,်, but =MMRLen(A2) would produce 2 only.
````VBA
'Requires MMRManipulator
'Just like Excel's Len function, this UDF returns the Length of a Myanmar Text String
Function MMRLen(target As Range) As Long
Dim targetLen As Long
    If target.Cells.CountLarge > 1 Then MMRLen = ">1Cell!": Exit Function
    targetLen = UBound(Split(MMRManipulator(target), "|")) + 1
    MMRLen = targetLen
End Function
````

#### 3.5.1.3.MMRLeft(target as Range, howMany as Long)as String
This works the same as Excel function Len but like the previous function, MMRLen, it works based on MMRLen rather than default function Len.
````VBA
'Requires MMRManipulator
'Just like Excel's Left function, this UDF extracts howMany number of words from a Myanmar Text String counted from the Left
Function MMRLeft(target As Range, howMany As Long) As String
Dim targetLen As Long
Dim MMRString As String
    If target.Cells.CountLarge > 1 Then MMRLeft = ">1Cell!": Exit Function
    If target.Value = "" Then MMRLeft = "": Exit Function
    targetLen = UBound(Split(MMRManipulator(target), "|")) + 1
    If targetLen > 0 Then
        If howMany = 0 Then MMRLeft = "": Exit Function
        If howMany >= targetLen Then MMRLeft = target.Value: Exit Function
        MMRString = MMRManipulator(target)
        MMRString = Application.WorksheetFunction.Substitute(MMRString, "|", "*|*", howMany)
        MMRString = Split(MMRString, "*|*")(0)
        MMRString = Replace(MMRString, "|", "")
        MMRLeft = MMRString
    Else
        If targetLen <= 0 Then MMRLeft = "": Exit Function
    End If
End Function
````

#### 3.5.1.4.MMRRight(target as Range, howMany as Long)as String
Same as previous function, MMRLeft, with the only difference being, from where we start counting just like the default function Right in Excel.
````VBA
'Requires MMRManipulator
'Just like Excel's Right function, this UDF extracts howMany number of words from a Myanmar Text String counted from the Right
Function MMRRight(target As Range, howMany As Long) As String
Dim targetLen As Long
Dim MMRString As String
    If target.Cells.CountLarge > 1 Then MMRRight = ">1Cell!": Exit Function
    If target.Value = "" Then MMRRight = "": Exit Function
    targetLen = UBound(Split(MMRManipulator(target), "|")) + 1
    If targetLen > 0 Then
        If howMany = 0 Then MMRRight = "": Exit Function
        If howMany >= targetLen Then MMRRight = target.Value: Exit Function
        MMRString = MMRManipulator(target)
        MMRString = Application.WorksheetFunction.Substitute(MMRString, "|", "*|*", targetLen - howMany)
        MMRString = Split(MMRString, "*|*")(1)
        MMRString = Replace(MMRString, "|", "")
        MMRRight = MMRString
    Else
        If targetLen <= 0 Then MMRRight = "": Exit Function
    End If
End Function
````

#### 3.5.1.5.MMRMid(target as Range, startPos as Long, howMany as Long)as String
This function, like the 2 above, was designed to behave just like Excel builtin function/formula, Mid. Be reminded that the counting was based on Myanmar word counting and not as English character counts.
````VBA
'Requires MMRManipulator
'Just like Excel's Mid function, this UDF extracts howMany number of words from a Myanmar Text String starting from startPos
Function MMRMid(target As Range, startPos As Long, howMany As Long) As String
Dim targetLen As Long
Dim upTill As Long
Dim MMRString As String
Dim lenArray()
    If target.Cells.CountLarge > 1 Then MMRMid = ">1Cell!": Exit Function
    If target.Value = "" Or startPos <= 0 Or howMany <= 0 Then MMRMid = "": Exit Function
    targetLen = UBound(Split(MMRManipulator(target), "|")) + 1
    If startPos > targetLen Then MMRMid = "": Exit Function
    If targetLen > 0 Then
        MMRString = MMRManipulator(target)
        upTill = IIf(startPos + howMany - 1 > targetLen, targetLen, startPos + howMany - 1)
        ReDim lenArray(startPos To upTill)
        lenArray = Evaluate("transpose(row(" & startPos & ":" & upTill & "))")
        MMRMid = Join(Application.Index(Split(MMRString, "|"), 0, lenArray), "")
    Else
        If targetLen <= 0 Then MMRMid = "": Exit Function
    End If
End Function
````

## 4.Supporting Formulas
A number of supporting basic formulas will be posted under this.\
These are just simple/basic formulas that users can edit/improved upon or replace with whatever they desired.\
For the sake of wider compatibility with different Excel versions, the following formulae are compatible with MS Excel versions as early as Office 2010 or maybe earlier.\
However, better formulas are being released with each new version of MS Office suite and some latest ones like FilterXML or SEQUENCE formula can be use if the user has higher version of Excel.\
The following are just to give the users an idea on how to extend the functionalities of the provided UDFs.
![MMRTokenizerXL](/images/formulas_for_wordcount_exploding.png)

## 5.Splitting/Exploding the tokenized words
MS Excel's builtin Text to Column function from Data tab in Menu doesn't see the UDF outputs in a cells as Text but rather like a formula!\
Therefore, it is advisable that the users should just copy paste the value from cells containing the UDFs over themselves or over to another column as values so that the Text-to-Column function can be used to split/explode the tokenized word output.

## 6.Releases
Releases can be found [here](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases).
1.  [v1.0a.First Release](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases/tag/v1.0a-Pre-Release) on 26NOV2021 19:40 Myanmar Standard Time.
2.  [v1.0.1a.Bugfix for Daw](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases/tag/v1.0.1a-Pre-Release) on 28NOV2021 07:51 Myanmar Standard Time.
3.  [v1.2.New functions](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases/tag/v1.2) on 20DEC2021 16:20 Myanmar Standard Time.

## 7.Proposed Uses
1. Sorting
2. Natural Language Processing
3. Transliteration/Transcription
4. Text-to-Speech/Speech-to-Text
5. Encryption
6. Word Count
7. Sentiment Analysis
8. Gender identification/prediction/analysis from salutation

## 8.The Future
I will probably write up another part here when I could successfully write an Excel formula based on the algorithm I used here.\
I believe that is very feasible but the only problem I can foresee now is that, Unichar and Unicode formulas are only available in Office 2013 onwards and unfortunately this will limit the users of my future tokenization formula.\
Fortunately for them, they can use the UDF that I developed, which may have a bit of a hassle for copying/importing the code into their worksheets.

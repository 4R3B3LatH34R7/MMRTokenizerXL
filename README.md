# MMRTokenizerXL
## Tokenizer for Myanmar Unicode Pyidaungsu Font (Visual Order)
![MMRTokenizerXL](/images/MMRTokenizerXL.png)
## The Journey
As usual, while I was haunting in my usual Facebook Excel groups, I came across somebody asking a question on how to sort a Myanmar Name in Myanmar font using the last name.
As we, all Myanmar people, know, our names do not have a first name/last name basis but rather like a jumble of nice words joined together to get a beautiful or following the Myanmar astrological beliefs.\
The main problem with this naming system concerns something else.

It is the way we write our language. We don't need to use a white space to separate each word or ?syllable. I am not a linguist. So, I don't know how it is called.\
Anyway, the real problem is that I found it hard to find out the word break points. I mean the exact point where one word of a name ends and where the next word starts.\
Let's consider a name in old Burmese e.g. ခင်မို့မို့အောင် (=Khin Moh Moh Aung) this is actually a Pali related name.\
In this case, even though we spell in English like so inside the parentheses, the Myanmar words do NOT have any white space between them. They just don't have to.\
If we want to separate/tokenize/split the name in English, we could easily do so in VBA using the SPLIT function or use some formulas in Excel worksheet UI.\
But we have no way of achieving the same goal with a Myanmar font, the Pyidaungsu font, namely, because we have no idea where the consonants are.

I am so interested in Natural Language Process (NLP) but it is very hard for me to understand those papers written by NLP specialists because I don't have the basics of linguistics and I was never officially trained in programming languages, especially the VBA, I have currently and mostly written in.\
However, I persevered and read through many papers but I got no where.\
Then I found some github pages where some people wrote up something about a python package and they showed an image which shows, in turn, a tokenized Myanmar words and I was so addicted to it.

Now that there is this question about how to sort Myanmar names using the last word of the name, I realized that it would be best to get the last consonant out and sort by that much like the English names do.\
So I tried to get the consonant of the last word from a name. It was very hard for me because I didn't know how Pyidaungsu font stitches each parts of a word.

For example, for the same name, ခင်မို့မို့အောင်, (a beautiful name for a girl btw), we actually typed it up using the Windows10's Burmese Keyboard (Visual Order) in Pyidaungsu font like (ခ, င, ် , မ, ိ, ု, ့ , မ, ိ, ု, ့ , ‌ေ, အ, ာ, င, ်     ) because we need to type the way we spell as well as the way we see/read it. There is another Burmese Keyboard using Phonetic Order but I am not using it and I don't develop code for that because the typing is pretty hard on it.\
From trial and error and trying to convert each typed parts into Unicode values using AscW and ChrW, I realized that, the developers of the Pyidaungsu font decided to put together a system (an incomplete one, more on this later) which, in a way, places the consonants in front of every other appendages of a Burmese word spelling.

What I meant is for the last part of the name, အောင်, they placed the consonant အ in front of the other stuff. And they bind them together like အ, ‌ေ, ာ, င, ် , and if that is a uniform method, I could come up with a UDF for tokenization!
So I tried to loop from the right-most part of the name and bam, the first obstacle hits me, right in the face.\
The င under the ် , the င is actually a consonant but it must be combined with the Athat or ် , to create the whole word. So I had to ignore the first consonant that comes after the Athat.\
The real problem now begins because Myanmar words have other part that comes between the Athat and the consonant!!! And the number of stuff that can come between is variable and can be more than 2 or 3.

I wrote incompleteness above because in some cases like, သန့်, we typed it up like သ, န, ် , ့   so I was thinking like the Athat ်   would come second last and the အောက်မြစ် ့  would come after it and I was totally wrong. The Athat ်  comes rightmost then the ့  comes after that which wreaks havoc in my algorithm and I had to spend hours trying to fix that.\
Therefore, I had to come up with a way to filter through to reach the consonant of the last word in a name. And I did.\
The algorithm is not graceful nor very cool but it works!

I was like OK, I can get the consonant of the last word in a name already. Why not get all the consonants using this algorighm.\
By this time, I can sort a column using the consonant of the last word of a name!\
I expanded and improved my code to be able extract all the consonants of a name!!!\
Then I found that using the consonant of the last word of a name for filtering is not very accurate because in Burmese language, we have a system of laying out the alphabetical order of words, which we called as, မြန်မာအက္ခရာစဉ်. This topic was taungt in the first year of my bachelor's degree and I failed this subject, the Myanmar language!\
I tried using the consonant extraction UDF and compared it to sorting by using only the consonant from the last word. And I found that the difference is quite impactful.\
This drives me to finish the whole Tokenization UDF because the Burmese language sorting system uses not just consonants but other parts in the building of a word like those I already mentioned above for example, ်  or ာ or ိ or ု or ့  etc of may of those things.\
Enough said and I wish to admit that I used just string manipulation functions in this UDF rather than the NLP methods. I am not employing ML or AI with all this.

## The UDFs
The source code of the UDFs may be released as plain text.\
There are currently 4 UDFs in the .bas and .xlsm files upon release.
1. [MMRTokenizer](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#11mmrtokenizer)
2. [MMRManipulator](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#12mmrmanipulator)
3. [getMMRConsonants](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#13getmmrconsonants)
4. [MMRParser](https://github.com/4R3B3LatH34R7/MMRTokenizerXL#14mmrparser)

### 1.1.MMRTokenizer
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

'Return a tokenized Myanmar String
Function MMRTokenizer(target As Range) As String
Dim ch As String
Dim returnString As String
Dim charCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
Dim previousCharAt As Integer '?long
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
                                If AscW(ch) = witecha Then
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

### 1.2.MMRManipulator
This is a tool spawned from being able tokenize in order to manipulate the Myanmar words typed in Pyidaungsu font.\
It can be used to tokenize the words, for any purpose, like for sorting, counting, replacing...etc...with the sky at the limit of users' imagination.\
With an argument switch, can change the tokenization character to become anything, any text, any string, even nothing!\
The users can also reverse the whole Myanmar sentence word by word with the first word becoming the last word and vice versa...\
The result of using this tool can be found in the [photo](/images/MMRTokenizerXL.png) under the right most column (Column F).
If cell A1 contains မိုးအောင်ခင်, then calling from inside cell B1 like, =MMRTokenizer(A1), shall return မိုး|အောင်|ခင်.

How we can use MMRManipulator to sort Myanmar names, words, sentences in reverse can be reviewed in the following short .gif.
![MMRManipulator](/images/MMRManipulator_sorting.gif)
The process is simple in that users just need to use the UDF to reverse the range containing Myanmar words.\
The UDF only requires the first argument, out of the existing 3: 
1. the target range, which is essential
2. the second argument is for defining the separator for the output of the UDF, which can be anything from ""=vbNullString or blank or a space or any word in Burmese or English and since no check was performed on this argument's validity, it can be quite powerful and dangerous at the same time.
3. the third argument is a boolean variable which acts as a switch for reversing the output of the UDF.

So, if cell A1 contains "ကိုကိုအေး" and from inside cell B1, if we call the UDF as: (let "->" denotes "returns")
1. =MMRManipulator(A1) -> ကို|ကို|အေး
2. =MMRManipulator(A1,"@") -> ကို@ကို@အေး
3. =MMRManipulator(A1,"",TRUE) -> အေးကိုကို (please note that "" is not space but denotes nothing)
Apart from the cell reference, the remaining 2 arguments are optional, thus, calling like =MMRManipulator(A1,,) is legitimate and will return ကို|ကို|အေး.

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

'Return tokenized words using user-selectable optional separator and ability to reverse the Myanmar word string
Function MMRManipulator(target As Range, Optional separator As String = "|", Optional reversed As Boolean = False) As String
Dim ch As String
Dim returnString As String
Dim charCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
Dim previousCharAt As Integer '?long
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
                                                   IIf(Len(returnString) > 0, separator, "") & _
                                                   IIf(reversed, Mid(target.Value, charCounter, previousCharAt - charCounter), returnString)
                                    previousCharAt = charCounter
                                Else
                                    previousChIsAthat = False
                                End If
                            Else
                                If AscW(ch) = witecha Then
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
    MMRManipulator = returnString
End Function
```

### 1.3.getMMRConsonants
This UDF was designed in the earlier stages of development of MMRTokenizer to help me identify, check and confirm the location of Myanmar consonants in a cell containing Myanmar word(s).\
There are altogether 4 possible arguments that can be passed when calling it.
1. target range (required)
2. reversed order (optional with default=false)
3. last character only (optional with default=false)
4. location of consonants (optional with default=false)
Apart from the target range, the rest as optional.\
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

'Return all consonants within a range with optional reversing, last character only or consonant locations instead of actual ones
Function getMMRConsonants(target As Range, Optional reversedOrder As Boolean = False, Optional lastCharOnly As Boolean = False, Optional LOC As Boolean = False) As String
Dim ch As String
Dim returnString As String
Dim charCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
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
                                If AscW(ch) = witecha Then
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

### 1.4.MMRParser
This UDF was also written in the earlier part of the development of MMRTokenizer to help me confirm the location of the Myanmar consonants.\
Only 1 argument is required, out of the 3 possible arguments.\
1. target range (required)
2. output as Myanmar (optional with default=false)
3. hightlight Consonants (optional with default=false)

This UDF just returns the Unicode values (numbers) as a string of text. For example, if cell A1 contains သီရိ, calling from inside cell B1 as,
1. =MMRParser(A1) -> 4126|4142|4123|4141 -> returning the Unicode values separated by | (pipe) character.
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

'Parse Myanmar strings into Unicode code values or Myanmar consonants and Diacritics and can return consonants in Burmese combined with numerical Diacritics
Function MMRParser(target As Range, Optional outputMMR As Boolean = False, Optional highlightConsonants As Boolean = False) As String
Dim returnStringArray()
Dim ch As String
Dim chCounter As Integer
Dim previousChIsAthat As Boolean
Dim shiftFfound As Boolean
Dim legitConsonantFound As Boolean
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
                                If AscW(ch) = witecha Then
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
## Formulas
A number of supporting basic formulas will be posted under this.\
These are just simple/basic formulas that users can edit/improved upon or replace with whatever they desired.\
For the sake of wider compatibility with different Excel versions, formulas which are compatible with as early as Office 2010 or earlier.\
However, better formulas are being released with each new version of MS Office suite like FilterXML or SEQUENCE formula can be use if the user has higher version of Excel.\
The following are just to give the users an idea on how to extend the functionalities of the provided UDFs.
![MMRTokenizerXL](/images/formulas_for_wordcount_exploding.png)

## Releases
Releases can be found [here](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases).
### v1.0a.First Release
[First release](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases/tag/v1.0a-Pre-Release) on 26NOV2021 19:40 Myanmar Standard Time.
### v1.0.1a.Bugfix for Daw
[Second release](https://github.com/4R3B3LatH34R7/MMRTokenizerXL/releases/tag/v1.0.1a-Pre-Release) on 28NOV2021 07:51 Myanmar Standard Time.

## The Future
I will probably write up another part here when I could successfully write an Excel formula based on the algorithm I used here.\
I believe that is very feasible but the only problem I can foresee now is that, Unichar and Unicode formulas are only available in Office 2013 onwards and unfortunately this will limit the users of my future tokenization formula.\
Fortunately for them, they can use the UDF that I developed, which may have a bit of a hassle for copying/importing the code into their worksheets.

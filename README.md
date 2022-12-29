# VBA Scripting Library
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/WolfpackWilson/VBA-Scripting-Library)](https://github.com/WolfpackWilson/VBA-Scripting-Library/releases)
[![GitHub All Releases](https://img.shields.io/github/downloads/WolfpackWilson/VBA-Scripting-Library/total)](https://github.com/WolfpackWilson/VBA-Scripting-Library/releases/latest)
[![GitHub issues](https://img.shields.io/github/issues/WolfpackWilson/VBA-Scripting-Library)](https://github.com/WolfpackWilson/VBA-Scripting-Library/issues)
[![GitHub](https://img.shields.io/github/license/WolfpackWilson/VBA-Scripting-Library)](https://github.com/WolfpackWilson/VBA-Scripting-Library/blob/master/LICENSE)

The goal of this repository is to recreate some of the scripting library included
on Windows OS so that programs can work on Mac OS, too.

## History
When running programs on the Mac OS, an error appears:<br>
`Error: Run-time error ’429’ ActiveX component can’t create object`

It appears that the 
[Mac OS doesn't have the Scripting Runtime Library](https://stackoverflow.com/questions/4670420/how-can-i-install-use-scripting-filesystemobject-in-excel-2010-for-mac). 
As such, anytime `CreateObject("Scripting.<object>")` is used, this error will appear on Mac OS.

## Installation and Use
Right click on the project in the projectect explorer then choose import file. Import any necessary files into your project. From there, call the objects to use them (i.e. `Dim dict As New Dictionary`).

### Similar Documentation
1. [ArrayList](https://excelmacromastery.com/vba-arraylist/)
    - Notable differences:
        - Use `arrList.Insert(Item, Pos)` instead of `arrList.Insert(Index, Value)`
1. [Dictionary](https://excelmacromastery.com/vba-dictionary/)
    - Notable differences:
        - Use `dict.Item(item)` instead of `dict(item)` when adding, editing, or viewing values.
        - Object keys are supported.
        - There is no case sensitivity option.

## Terms of Service
As defined in the [MIT License](https://github.com/TheEric960/VBA-Scripting-Library/blob/master/LICENSE):
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Version 2.6	30/5/00


Extract all files in the Zip archive with the 'use folder names' box checked

Although this program was programmed in VB6.0 SP3 it will probably work with versions 4 upwards

I have only commented the complicated bits (and not very well) so i recommend this code shouldn't be read by beginners
because although the basic idea (as outlined in the text file) is simple but the code to put it into practice is very complicated!

To change the depth press on the number below the 'Max Ply Depth' label then press enter

Any help with improving this code optimizing it or debugging it would be greatly appreciated!
 
Please send any improvements in the code in a zip file to dioxic@madasafish.com

If you do provide any help your name will be put in the credits list below

Thanx Mark!


Credits:

Main Program		Mark Baker-Munton
Graphics		Mark Baker-Munton
Initial Idea		Daniel George
Bug Fix			KD
Bug Fix			RJ (rwj@post.com)
Bug Fix			Ulli (umgedv@aol.com)
Bug Fix			John Pettit (masteryoda@www.com)
Checkers Rules		Ulli (umgedv@aol.com)


Versions:

2.6	Added - Full International checkers rules implemented!! (read rules.txt for details of the rules)
	Fixed - Being able to take pieces when there are again the side

2.5	Added - Simple Alpha Beta Pruning routine

2.4	Fixed - Being able to take without moving or being a double a piece diagonally behind
	Fixed - The board is now correct (bottom right square is now white) as stated by the rules

2.3	Fixed - Occasional locked array error workaround implemented

2.2	Fixed -	not being able to take multiple pieces in single player mode

2.1	Fixed -	starting with depth of 0 in first game causes stalemate to occur

2.0	Info -	second generation AI Engine first implemented
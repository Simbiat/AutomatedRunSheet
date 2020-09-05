# What?
Excel-based run sheet (check list) enhanced with VBA for tracking and automation capabilities. That is you create a set of tasks (or steps per task) that you do daily and then mark as completed.

# Where and when?
It's best to be used in collaborative environments, where a set of tasks are done simultaneously by several people. Of course, if you can utilize a proper database and interface written relying on it - use that, but in some cases that may be impossible, too expensive, or too complicated.

# Why?
Above partially answers this question already: to track completion of daily tasks. Some workflows may require manual or semi-automatic processing done several people on regular basis. If it's just 1 or 2 tasks - that's easy. But if you have several dozens of operations and more than just one person - that may be problematic.

# How?
It's relatively simple, really. You can download Sample ARS.xlsm file and fill it in with your tasks. It has integrated `EditorManual`, that gives some explanation on what's what inside there and a bunch of settings you may want to utilize. Generally, it's expected that you will go to `VBA Editor` only to change password for the book: this should be done right at the top of `Security` module, you won't miss it.

# Features
* Marking steps status with either drop-down menu or shortkeys, that can be customized in `Settings`
* Turning steps into buttons allowing running of functions, including checking for files, sending mails, starting programs, taking screenshots
* Marking of steps, that have time limit (needs to be started at or be finished by specific time)
* `End of day` function, that will create backup of the sheet (protected by randomly generated password), clean current one and change current _business_ date
* Making steps specific for `Late processing` (for cases, when you are time limited, but limit failed), that will be shown only if `Late mode` is enabled
* Each step can be hidden during special processing days (like working day during public holiday) using appropriate setting for each step
* Extensive logging (sample available in repository)
* Substantial customization through `Settings` sheet, including colors of buttons and statuses of steps
* Easy to use `Editor mode`, that can be assigned to specific users only through `Settings`
* When not in `Editor mode`, workbook is shared and protected from editing as much as possible in Excel (using password, that needs to be edited through VBA Editor)

# Customization
The sheet is pretty customizable already through various settings, that have explanation on what they mean, although has some restrictions, like names of worksheets, placement of general elements. It has some pre-built functions you may want to use in some cases, although they are based on what we need to use at my work. In case you want to add some custom functions - you can do that, there is a logic supporting custom ones for buttons. If you think that some _general_ function can be useful - create an _Issue_ and I will consider adding it. Or make a _pull request_ with it, even.

If you want something built specifically for your company that is also possible, obviously. If you want that, but do not know how - contact me at simbiat@outlook.com with details of what you need. I will check what is your demand and let you know my price for coding that. Full code of the solution will not be published in order to protect your company, but my main condition in this endeavor will be possibility to share its pieces (that is sub-functions), that may be useful to people outside of your company.

# History
At my work we have hundreds of steps done by 2 to 4 people at the same time. Long ago we used a printed sheet where respective tasks were marked as completed by signatures chosen by each person. When a demand to also write down time came from management, it was obvious, that something needed to be changed, so instead of printing Excel sheet, I've transformed it into a macro-driven application to do just that and over the years it got some useful functions, that simplified operators' work. It was not perfect, though, since it almost always required me to modify the sheet. Thus, I decided to rewrite it from scratch, utilize my current knowledge of how to write not just a script, but a "product", that can be utilized with little to know knowledge of VBA. This is what it resulted in.

# Future
I do not plan on _actively_ developing this further. If any requests (issues or pulls or whatnot) come I will address them at best effort basis. That includes bugs, obviously.

# Help and support
If details in `EditorManual` are not enough and you need some more guidance, please, raise an _Issue_ to this repository and I will try to address it both in the issue and in the manual. Same for any bugs or improvements.
And there will be bugs, most likely. Although I tested all of the functions, I can't test the workbook completely in _Shared_ environment, so something else will pop-up. And there are also some behaviors, that I already know about, but not sure how to cure (they even may be limitations of Excel itself):
1. Sometimes validation fails to be applied to a cell. I've added a catch for that to show warning and suggest re-opening the workbook, but to this day I do not understand what exactly causes it with only an assumption, that it's related to book being shared through OneDrive or SharePoint.
2. At the moment when removing several buttons and regular steps at the same time `Method 'Range' of object '_Global failed` error can be encountered, but it's not always there. I am unable to see a sure-fire way to trigger it, and thus I do not see where exactly it needs to be fixed.

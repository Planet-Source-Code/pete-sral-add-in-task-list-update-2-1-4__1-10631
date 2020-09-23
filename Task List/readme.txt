PjS - Task List Add-In Update 2.1.4
4:29 PM 08/24/2000

- Modified: Set version for task on completion rather than on entry.
- Fixed: New VB Project addin is disabled until VB project is saved to file.

PjS - Task List Add-In Update 2.1
5:47 PM 7/25/00,11:16 PM 08/09/2000,8:07 AM 08/11/2000

- Added: Column Priority
- Added: Column Version (Default: hidden)
- Modified: Column Added is now hidden (Default: hidden)
- Added: New "Add Task" form
- Modified: Double Click on TaskList will display menu instead of adding.
- Added: Edit Task menu option (allows to edit the description only).
- Added: Icons to Column headers to show current sort when clicked on.
- Added: Save column widths 
- Added: Small toolbar on the left hand side
- Added: Select Grid (alternating) colors

B. Harriger Task List Update 2.0
24 Jul 2000

- Modified: Cleaned up coding convention & formatting a little
- Added:  Project Group capability.  Now works properly with projects
          in project groups.
- Added:  Column sort.  Click on column header to toggle asc/desc
          sorts.  Tasks are saved in sorted order.
- Still Broken:  The toolbar button isn't quite right yet.



PjS Task List Add-In Update
By: Pete Sral

5:57 PM 7/16/00
- Added: Add and Completed Date Columns and Visible Columns headers
- Added: Error Handling Code
- Added: About Form (Right Click)
- Added: View Form - for cut and paste purposes, Print feature
- Added: Delete from Grid ( Right Click)

Coming Soon....
- Bug History Addin


--- Original Author ----
VB Task List Add-In
By: Mark Joyal

I created this VB add-in to help me out when coding projects.  I wanted 
something to allow me to add a todo list right inside the VB IDE.  This 
addin shows a listview of tasks with checkboxes for each.  The window 
will dock to the IDE, and you can add new tasks by right clicking on the
listview and choosing "new task" or by doubleclicking in an empty slot.
You can edit tasks directly in the listview via the 2 click method.
Completed items are shown in gray(disabled) text.  A different task list
is kept for each project.  note: multiple projects open in the same 
instance of VB is not supported...yet...There is alot of good stuff here,
I hope you like it.

If you wish to use just the add-in and not bother with the code, just 
copy the .dll file to your windows/system32 directory and run regsvr32 
on it.

Mark Joyal
mark@thejoyals.net
http://www.thejoyals.net/Sourcecode

#' Get the Repository object
dim repository
set repository = CreateObject("EA.Repository")

' Get the current diagram
dim currentDiagram
set currentDiagram = repository.GetCurrentDiagram()

' Check if a diagram is open
if not currentDiagram is nothing then
' Activate the diagram tab
repository.ShowInProjectView currentDiagram
else
' No diagram is open, handle accordingly
Session.Output "No diagram is currently open."
end if

' Release the Repository object
set repository = nothing

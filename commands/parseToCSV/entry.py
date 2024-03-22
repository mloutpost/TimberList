import adsk.core
import os
from ...lib import fusion360utils as futil
from ... import config
import adsk.fusion
import traceback
import csv
from fractions import Fraction

app = adsk.core.Application.get()
ui = app.userInterface


# TODO *** Specify the command identity information. ***
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_cmdDialog'
CMD_NAME = 'Timber List'
CMD_Description = 'Create csv with timber data from selected timbers.'

# Specify that the command will be promoted to the panel.
IS_PROMOTED = True

# TODO *** Define the location where the command button will be created. ***
# This is done by specifying the workspace, the tab, and the panel, and the 
# command it will be inserted beside. Not providing the command to position it
# will insert it at the end.
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidScriptsAddinsPanel'
COMMAND_BESIDE_ID = 'Timber List'

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []

# Executed when add-in is run.
def start():
    # Create a command Definition.
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    # Define an event handler for the command created event. It will be called when the button is clicked.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******** Add a button into the UI so the user can run the command. ********
    # Get the target workspace the button will be created in.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get the panel the button will be created in.
    panel = workspace.toolbarPanels.itemById(PANEL_ID)

    # Create the button command control in the UI after the specified existing command.
    control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)

    # Specify if the command is promoted to the main toolbar. 
    control.isPromoted = IS_PROMOTED


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()


# Function that is called when a user clicks the corresponding button in the UI.
# This defines the contents of the command dialog and connects to the command related events.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Created Event')

    # https://help.autodesk.com/view/fusion360/ENU/?contextId=CommandInputs
    inputs = args.command.commandInputs

    # TODO Define the dialog for your command by adding different inputs to the command.

    selectionInput = inputs.addSelectionInput(CMD_ID + '_selection', 'Select',
                                              'Select bodies or occurrences')
    selectionInput.setSelectionLimits(1)

    # TODO Connect to the events that are needed by this command.
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(args.command.validateInputs, command_validate_input, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)


# This event handler is called when the user clicks the OK button in the command dialog or 
# is immediately called after the created event not command inputs were created for the dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Execute Event')

    inputs = args.command.commandInputs
    selection: adsk.core.SelectionCommandInput = inputs.itemById(CMD_ID + '_selection')

    # TODO ******************************** Your code here ********************************

    # Get a reference to your command's inputs.
    futil.log(f'Inputs: {inputs}')

    objects = getSelectedObjects(selection)
    obj_properties = {}

    for obj in objects:
        qty = obj.sourceComponent.allOccurrencesByComponent(obj.component).count
        obj_properties[obj.component.name] = [TimberData(obj).timberProperties(),
                                              qty]  # uses dictionary to eliminate duplicate occurances in the output and obtain count

    # Do something interesting
    fileDialog = ui.createFileDialog()
    fileDialog.isMultiSelectEnabled = False
    fileDialog.title = "filename"
    fileDialog.filter = 'CSV (*.csv);;TXT (*.txt);;All Files (*.*)'
    fileDialog.filterIndex = 0
    dialogResult = fileDialog.showSave()
    if dialogResult == adsk.core.DialogResults.DialogOK:
        filename = fileDialog.filename
    else:
        return
    with open(filename, 'w', newline='') as csvfile:
        fieldnames = ['Name', 'Qty', 'Length', 'Width', 'Height']
        writer = csv.writer(csvfile)
        writer.writerow(fieldnames)
        for name, obj in obj_properties.items():
            row = [name, obj[1], obj[0]['length'], obj[0]['width'], obj[0]['height']]
            writer.writerow(row)


# This event handler is called when the command needs to compute a new preview in the graphics window.
def command_preview(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Preview Event')
    inputs = args.command.commandInputs


# This event handler is called when the user changes anything in the command dialog
# allowing you to modify values of other inputs based on that change.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.inputs

    # General logging for debug.
    futil.log(f'{CMD_NAME} Input Changed Event fired from a change to {changed_input.id}')


# This event handler is called when the user interacts with any of the inputs in the dialog
# which allows you to verify that all of the inputs are valid and enables the OK button.
def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Validate Input Event')

    inputs = args.inputs

    args.areInputsValid = True

    # Verify the validity of the input values. This controls if the OK button is enabled or not.
    #valueInput = inputs.itemById('value_input')
    #if valueInput.value >= 0:
    #    args.areInputsValid = True
    #else:
    #    args.areInputsValid = False
        

# This event handler is called when the command terminates.
def command_destroy(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Destroy Event')

    global local_handlers
    local_handlers = []

### MY CODE

def getSelectedObjects(selectionInput):
    objects = []
    for i in range(0, selectionInput.selectionCount):
        selection = selectionInput.selection(i)
        selectedObj = selection.entity
        if type(selectedObj) is adsk.fusion.Occurrence:
            objects.append(selectedObj)
    return objects


def dec_to_proper_frac(dec):
    sign = "-" if dec < 0 else ""
    frac = Fraction(abs(dec))
    if frac.numerator % frac.denominator == 0:
        output = f"{sign}{frac.numerator // frac.denominator}"
    else:
        output = (f"{sign}{frac.numerator // frac.denominator} "
                  f"{frac.numerator % frac.denominator}/{frac.denominator}")
    return output


def roundPartial(value, resolution):
    return round(value / resolution) * resolution


class TimberData:
    def __init__(self, fusionObject):
        self.fusionObject = fusionObject

    def timberProperties(self):
        sel_prop = {}

        if type(self.fusionObject) is adsk.fusion.BRepBody or \
                type(self.fusionObject) is adsk.fusion.Occurrence:
            min_box = self.fusionObject.orientedMinimumBoundingBox
            dimensions = [min_box.length, min_box.width, min_box.height] # access raw output from minimum bounding box object, names don't matter yet
            dim_sorted = sorted(dimensions, reverse=True)
            length, width, height = str(dec_to_proper_frac(roundPartial((dim_sorted[0]) / 2.54, 0.25))) + '"', \
                                    str(dec_to_proper_frac(roundPartial(dim_sorted[1] / 2.54, 0.25))) + '"', \
                                    str(dec_to_proper_frac(roundPartial(dim_sorted[2] / 2.54, 0.25))) + '"'
            sel_prop["Occurrence"] = self.fusionObject.name
            sel_prop["length"], sel_prop["width"], sel_prop["height"] = length, width, height
        return sel_prop

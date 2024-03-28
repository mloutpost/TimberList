import adsk.core
import os
from ...lib import fusion360utils as futil
from ... import config
import adsk.fusion
import traceback
import csv
from fractions import Fraction
import math

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
TOOLBAR_TAB = 'SolidTab'
PANEL_ID = 'AssemblePanel'
COMMAND_BESIDE_ID = 'Timber List'

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []

wood_species = {
        'Model': [True, 1.0],
        'Ash, Black': [False, 0.833],
        'Ash, Green': [False, 0.849],
        'Ash, White': [False, 0.769],
        'Basswood': [False, 0.673],
        'Beech': [False, 0.865],
        'Black Locust': [False, 0.929],
        'Black Walnut': [False, 0.913],
        'Bur Oak': [False, 0.993],
        'Cedar, Alaska': [False, 0.577],
        'Cedar, Eastern Red': [False, 0.593],
        'Cedar, Northern White': [False, 0.449],
        'Cedar, Southern White': [False, 0.416],
        'Cedar, Western Red': [False, 0.432],
        'Cherry, Black': [False, 0.721],
        'Chestnut': [False, 0.881],
        'Cypress, Southern': [False, 0.817],
        'Douglas Fir, Coast Region': [False, 0.609],
        'Douglas Fir, Rocky Mountains': [False, 0.561],
        'Elm': [False, 0.865],
        'Elm, red': [False, 0.785],
        'Elm, white': [False, 0.881],
        'Fir, Balsam': [False, 0.721],
        'Fir, Commercial White': [False, 0.737],
        'Gum, Black': [False, 0.721],
        'Gum, Red': [False, 0.801],
        'Hackberry': [False, 0.817],
        'Hemlock, Eastern': [False, 0.801],
        'Hemlock, Western': [False, 0.657],
        'Hickory': [False, 1.025],
        'Hickory, Pecan': [False, 0.993],
        'Honeylocust': [False, 0.929],
        'Larch': [False, 0.769],
        'Locust': [False, 0.929],
        'Maple, Bigleaf': [False, 0.753],
        'Maple, Black': [False, 0.865],
        'Maple, Red': [False, 0.801],
        'Maple, Silver': [False, 0.721],
        'Maple, Soft': [False, 0.801],
        'Maple, Sugar': [False, 0.897],
        'Mulberry': [False, 0.945],
        'Oak, Post': [False, 1.025],
        'Oak, Red': [False, 0.977],
        'Oak, White': [False, 1.009],
        'Osage Orange': [False, 1.025],
        'Pecan': [False, 0.993],
        'Pine, Lodgepole': [False, 0.625],
        'Pine, Northern white': [False, 0.577],
        'Pine, Norway': [False, 0.673],
        'Pine, Ponderosa': [False, 0.721],
        'Pine, Southern Yellow': [False, 0.849],
        'Pine, Sugar': [False, 0.833],
        'Poplar, Yellow': [False, 0.609],
        'Redwood, American': [False, 0.801],
        'Spruce, Canadian': [False, 0.545],
        'Spruce, Engelman': [False, 0.625],
        'Spruce, Sitka': [False, 0.529],
        'Sycamore': [False, 1.009],
        'Tamarack': [False, 0.753],
        'Willow': [False, 0.865]
    }

# Executed when add-in is run.
def start():
    # Create a command Definition.
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    # Define an event handler for the command created event. It will be called when the button is clicked.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******** Add a button into the UI so the user can run the command. ********
    # Get the target workspace the button will be created in.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get the tab the button will be created in.
    solidTab = workspace.toolbarTabs.itemById('SolidTab')

    # Get the panel the button will be created in.
    panel = solidTab.toolbarPanels.itemById(PANEL_ID)

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

    # Add first user entry
    selectionInput = inputs.addSelectionInput(CMD_ID + '_selection', 'Timbers',
                                              'Select timbers to add to CSV Timber List')  # returns object(s) from selection
    selectionInput.setSelectionLimits(1)  # set limit to >= 1
    selectionInput.addSelectionFilter(adsk.core.SelectionCommandInput.Occurrences)  #  Basically limit selection to components

    # Add second user entry
    inputs.addTextBoxCommandInput(CMD_ID + '_partPrefix', 'Part Number Prefix', "LCTF-", 1, False)

    # Add third user entry
    dropdownInput = inputs.addDropDownCommandInput(CMD_ID + "_species", 'Wood Species',
                                                             adsk.core.DropDownStyles.TextListDropDownStyle)
    dropdownItems = dropdownInput.listItems
    for name, status in wood_species.items():
        dropdownItems.add(name, status[0])


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

    #  creates objects from our above user inputs so we can pull data from them
    selection: adsk.core.SelectionCommandInput = inputs.itemById(CMD_ID + '_selection')
    partPrefix: adsk.core.TextBoxCommandInput = inputs.itemById(CMD_ID + '_partPrefix')
    speciesData: adsk.core.DropDownCommandInput = inputs.itemById(CMD_ID + '_species')


    # TODO ******************************** Your code here ********************************

    # Get a reference to your command's inputs.
    futil.log(f'Inputs: {inputs}')

    objects = getSelectedObjects(selection)  #  Calls separate function to validate and return components from selection.
    obj_properties = {}  # create a dictionary to hold all the properties
    part_index = 1  # start the part index at 1

    # uses dictionary to eliminate duplicate occurrences in the output and obtain count
    for obj in objects:
        qty = obj.sourceComponent.allOccurrencesByComponent(obj.component).count

        # Dictionary contains the following: component name, list of bounding box properties, qty of occurrences, mass, material
        obj_properties[obj.component.name] = [TimberData(obj).timberProperties(),
                                                qty, TimberData(obj).getMass(speciesData),
                                                TimberData(obj).getMaterial(speciesData)]

    for key in obj_properties.keys():  # Create part numbers iterating over dictionary to avoid duplicates
        part_number = str(partPrefix.text) + str(part_index)
        obj_properties[key].append(part_number)
        part_index += 1


    # Do something interesting
    fileDialog = ui.createFileDialog()
    fileDialog.isMultiSelectEnabled = False
    fileDialog.title = "filename"
    fileDialog.filter = 'CSV (*.csv)'
    fileDialog.filterIndex = 0
    dialogResult = fileDialog.showSave()
    if dialogResult == adsk.core.DialogResults.DialogOK:
        filename = fileDialog.filename
    else:
        return

    # Writes CSV with all the collected fields
    with open(filename, 'w', newline='') as csvfile:
        fieldnames = ['Name', 'Part #', 'Material', 'Qty', 'Order Length - ft', 'Order Width - in', 'Order Height - in',
                      "Total Boardfeet", "Order Mass - kg", "Exact Length - in", "Exact Width - in", "Exact Height - in",
                      'Exact Mass - kg']
        writer = csv.writer(csvfile)
        writer.writerow(["Length field is rounded up to the nearest even, and 2' is added for ordering purposes."])
        writer.writerow(fieldnames)
        for name, obj in obj_properties.items():
            row = [name, obj[4], obj[3], obj[1], obj[0]['length'], obj[0]['width'], obj[0]['height'],
                   str(float(obj[0]["boardFeet"])*float(obj[1])), "placeholder", obj[0]["r_length"], obj[0]["r_width"],
                   obj[0]["r_height"], obj[2]]
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
    dropdownInput: adsk.core.DropDownCommandInput = inputs.itemById(CMD_ID + '_species')
    futil.log(f'Selected: {dropdownInput.selectedItem.name}')


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
    '''Builds a list of components. Contains duplicates on return.'''
    objects = []
    for i in range(0, selectionInput.selectionCount):
        selection = selectionInput.selection(i)
        selectedObj = selection.entity
        if type(selectedObj) is adsk.fusion.Occurrence:
            objects.append(selectedObj)
    return objects


def dec_to_proper_frac(dec):
    '''Float to arch notation for ordering.'''
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

    """Where a timber is a component... this is the backbone of the add-in. Uses built in fusion command
     "MinimumBoundingBox" to generate a tight box around the body. This creates the smallest size timber
      necessary for complex curves, which saves money during ordering. This class also can process and store
      data for other component properties."""

    def __init__(self, fusionObject):
        self.fusionObject = fusionObject

    def timberProperties(self):
        '''Takes the input of a fusion object (component) and returns dimensions for a tight bounding box, and volume
        measured in boardfeet. Units are also converted to arch style.'''
        sel_prop = {}

        if type(self.fusionObject) is adsk.fusion.BRepBody or \
                type(self.fusionObject) is adsk.fusion.Occurrence:
            min_box = self.fusionObject.orientedMinimumBoundingBox
            dimensions = [min_box.length, min_box.width, min_box.height] # access raw output from minimum bounding box object, names don't matter yet
            dim_sorted = sorted(dimensions, reverse=True)
            length, width, height = ((dim_sorted[0]/12) / 2.54), \
                                    str(dec_to_proper_frac(roundPartial(dim_sorted[1] / 2.54, 0.125))), \
                                    str(dec_to_proper_frac(roundPartial(dim_sorted[2] / 2.54, 0.125))) # length value rounds to nearest foot
            real_length, real_width, real_height = str(dec_to_proper_frac(roundPartial((dim_sorted[0]) / 2.54, .125))), \
                                    str(dec_to_proper_frac(roundPartial(dim_sorted[1] / 2.54, 0.125))), \
                                    str(dec_to_proper_frac(roundPartial(dim_sorted[2] / 2.54, 0.125)))
            rounded_length = roundEven(length) + 2  # rounds to nearest 2 and then adds 2'
            board_feet = rounded_length*(dim_sorted[1]/2.54)*((dim_sorted[2]/2.54)/12)
            sel_prop["Occurrence"] = self.fusionObject.name
            sel_prop["length"], sel_prop["width"], sel_prop["height"], sel_prop["boardFeet"], sel_prop["r_length"], \
                sel_prop["r_width"], sel_prop["r_height"] = \
                rounded_length, width, height, round(board_feet), real_length, real_width, real_height
        return sel_prop

    def getMass(self, species_data):
        '''Requires command dropdown input from 'Command Created' function. If the default 'Model' parameter
        is selected, it uses the default mass from fusion, applied individually to each timber. This is set by
        the material assigned in Fusion. Otherwise, it references a dictionary of known densities for green lumber
        and multiplies this by the volume of the object to produce the mass.'''
        if species_data.selectedItem.name == 'Model':
            return round(self.fusionObject.physicalProperties.mass, 1)
        else:
            mass = wood_species[species_data.selectedItem.name][1] * self.fusionObject.physicalProperties.volume
            return round(mass/1000, 1)  # convert to kg

    def getMaterial(self, species_data):
        """Uses drop down input to return material type on csv output."""
        if species_data.selectedItem.name == 'Model':
            return self.fusionObject.getPhysicalProperties
        else:
            return species_data.selectedItem.name


def roundEven(f):
    return math.ceil(f / 2.) * 2

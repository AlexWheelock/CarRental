Option Explicit On
Option Strict On
Option Compare Binary
'Alex Wheelock
'RCET 0625
'Car Rental
'Spring 2024
'https://github.com/AlexWheelock/CarRental

Imports System.Security.Cryptography.X509Certificates

Public Class RentalForm

    'Sets the defaults, but does not reset the summary
    'Does not reset the summary button because once one customers transaction is completed,
    'we can continue to look at the summary data
    Sub Defaults()
        NameTextBox.Text = ""
        NameTextBox.BackColor = Color.White
        AddressTextBox.Text = ""
        AddressTextBox.BackColor = Color.White
        CityTextBox.Text = ""
        CityTextBox.BackColor = Color.White
        StateTextBox.Text = ""
        StateTextBox.BackColor = Color.White
        ZipCodeTextBox.Text = ""
        ZipCodeTextBox.BackColor = Color.White
        BeginOdometerTextBox.Text = ""
        BeginOdometerTextBox.BackColor = Color.White
        EndOdometerTextBox.Text = ""
        EndOdometerTextBox.BackColor = Color.White
        DaysTextBox.Text = ""
        DaysTextBox.BackColor = Color.White
        MilesradioButton.Checked = True
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub

    'Validates user inputs
    Function ValidateInputs() As Boolean
        Dim valid As Boolean = True
        Dim validateString As Integer
        Dim validateAddress() As String = Split(AddressTextBox.Text, " ")
        Dim validateNumber As Integer
        Dim beginningMiles As Integer
        Dim endingMiles As Integer
        Dim errorMessage As String = ("The following information is incorrect:" & vbCrLf _
            & vbCrLf)

        'Validates that a name has been entered, and that the name is not a number
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 

        'If there is an input that can be converted to an integer, then it is a number and invalid
        Try
            validateString = CInt(NameTextBox.Text)
            valid = False
            NameTextBox.Focus()
            NameTextBox.BackColor = Color.LightYellow
            NameTextBox.Text = ""
            errorMessage += "Name cannot contain a number"
        Catch ex As Exception
            'if the input name is not a number, then checks to ensure that something was entered
            If NameTextBox.Text = "" Then
                valid = False
                NameTextBox.Focus()
                NameTextBox.BackColor = Color.LightYellow
                errorMessage += "Please enter a name"
            Else
                NameTextBox.BackColor = Color.White
            End If
        End Try

        'Checks that an address was entered, with both a home number of a whole number of at least 1, and a street name that is not a number
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 

        'Checks to ensure that something was entered into the AddressTextBox
        If AddressTextBox.Text = "" Then
            If valid Then
                errorMessage += "Please enter an address"
                AddressTextBox.Focus()
            Else
                errorMessage += ", please enter an address"
            End If
            valid = False
            AddressTextBox.BackColor = Color.LightYellow
        Else
            'If something was entered, then it checks that the format is correct
            'AddressTextBox is put into an array, and if the index 0 is not a number then it fails
            'number must also be at least 1
            Try
                validateNumber = CInt(validateAddress(0))
                If CInt(validateAddress(0)) > 0 Then
                    AddressTextBox.BackColor = Color.White
                Else
                    'home number was less than 1
                    If valid Then
                        errorMessage += "The home number must be at least 1"
                        AddressTextBox.Focus()
                    Else
                        errorMessage += ", the home number must be at least 1"
                    End If
                    valid = False
                    AddressTextBox.Text = ""
                    AddressTextBox.BackColor = Color.LightYellow
                End If
            Catch ex As Exception
                'entered address number is not a number
                If valid Then
                    errorMessage += "The address must contain a home number"
                    AddressTextBox.Focus()
                Else
                    errorMessage += ", the address must contain a home number"
                End If
                valid = False
                AddressTextBox.Text = ""
                AddressTextBox.BackColor = Color.LightYellow
            End Try

            'Then checks index 1 to ensure that a street name is entered and is not a number
            Try
                'If index 1 exists at all, it checks it's contents and validates it
                If validateAddress(1) = "" Then
                    If valid Then
                        errorMessage += "The address must contain a street name"
                        AddressTextBox.Focus()
                    Else
                        errorMessage += ", the address must contain a street name"
                    End If
                    valid = False
                    AddressTextBox.Text = ""
                    AddressTextBox.BackColor = Color.LightYellow
                Else
                    AddressTextBox.BackColor = Color.White
                End If
            Catch ex As IndexOutOfRangeException
                'If index 1 does not exist, then it tells the user to enter a street name
                'and the validation process fails
                If valid Then
                    errorMessage += "The address must contain a street name"
                    AddressTextBox.Focus()
                Else
                    errorMessage += ", the address must contain a street name"
                End If
                valid = False
                AddressTextBox.Text = ""
                AddressTextBox.BackColor = Color.LightYellow
            End Try
        End If

        'Validates that a city was entered, and that the city is not a number
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 
        If CityTextBox.Text = "" Then
            If valid Then
                errorMessage += "Please enter a city"
                CityTextBox.Focus()
            Else
                errorMessage += ", please enter a city"
            End If
            valid = False
            CityTextBox.BackColor = Color.LightYellow
        Else
            'Something was entered into CityTextBox and now needs to be checked whether or not it is a number
            Try
                validateString = CInt(CityTextBox.Text)
                'city was a number and is invalid
                If valid Then
                    errorMessage += "The city cannot be a number"
                    CityTextBox.Focus()
                Else
                    errorMessage += ", the city cannot be a number"
                End If
                valid = False
                CityTextBox.Text = ""
                CityTextBox.BackColor = Color.LightYellow
            Catch ex As Exception
                'city is valid
                CityTextBox.BackColor = Color.White
            End Try
        End If

        'Validates that a state was entered and that the state is not a number
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 
        If StateTextBox.Text = "" Then
            If valid Then
                errorMessage += "Please enter a state"
                StateTextBox.Focus()
            Else
                errorMessage += ", please enter a state"
            End If
            valid = False
            StateTextBox.BackColor = Color.LightYellow
        Else
            'Something was entered into the StateTextBox and now needs to be verified that it is not a number
            Try
                validateString = CInt(StateTextBox.Text)
                If valid Then
                    errorMessage += "The state cannot be a number"
                    StateTextBox.Focus()
                Else
                    errorMessage += ", the state cannot be a number"
                End If
                valid = False
                StateTextBox.Text = ""
                StateTextBox.BackColor = Color.LightYellow
            Catch ex As Exception
                'state is valid
                StateTextBox.BackColor = Color.White
            End Try
        End If

        'Validate that a zip code was entered, and that it is at least 5 numbers
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 
        If ZipCodeTextBox.Text = "" Then
            If valid Then
                errorMessage += "Please enter a zip code"
                ZipCodeTextBox.Focus()
            Else
                errorMessage += ", please enter a zip code"
            End If
            valid = False
            ZipCodeTextBox.BackColor = Color.LightYellow
        Else
            'Something was entered into ZipCodeTextBox and now needs to be validated
            Try
                validateNumber = CInt(ZipCodeTextBox.Text)
                ZipCodeTextBox.BackColor = Color.White
                'verified that it is a number, and now needs to be at least 5 digits
                If Len(ZipCodeTextBox.Text) < 5 Then
                    If valid Then
                        errorMessage += "The zip code must be at least 5 digits"
                        ZipCodeTextBox.Focus()
                    Else
                        errorMessage += ", the zip code must be at least 5 digits"
                    End If
                    valid = False
                    ZipCodeTextBox.Text = ""
                    ZipCodeTextBox.BackColor = Color.LightYellow
                Else
                    If CInt(ZipCodeTextBox.Text) >= 0 Then
                    Else
                        'input zip code was negative and invalid
                        If valid Then
                            errorMessage += "The zip code cannot be negative"
                            ZipCodeTextBox.Focus()
                        Else
                            errorMessage += ", the zip code cannot be negative"
                        End If
                        valid = False
                        ZipCodeTextBox.Text = ""
                        ZipCodeTextBox.BackColor = Color.LightYellow
                    End If
                End If
            Catch ex As Exception
                'validation failed and is not a number
                If valid Then
                    errorMessage += "The zip code must be a number"
                    ZipCodeTextBox.Focus()
                Else
                    errorMessage += ", the zip code must be a number"
                End If
                valid = False
                ZipCodeTextBox.Text = ""
                ZipCodeTextBox.BackColor = Color.LightYellow
            End Try
        End If

        'Validates that the beginning odometer miles was entered and is a number that is not negative
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 
        If BeginOdometerTextBox.Text = "" Then
            If valid Then
                errorMessage += "Please enter the beginning odometer miles"
                BeginOdometerTextBox.Focus()
            Else
                errorMessage += ", please enter the beginning odometer miles"
            End If
            valid = False
            BeginOdometerTextBox.BackColor = Color.LightYellow
        Else
            'Something was entered into BeginOdometerTextBox and now needs to be verified as a number that is not negative
            Try
                validateNumber = CInt(BeginOdometerTextBox.Text)
                beginningMiles = CInt(BeginOdometerTextBox.Text)
                If beginningMiles >= 0 Then
                    BeginOdometerTextBox.BackColor = Color.White
                Else
                    'input number was negative
                    If valid Then
                        errorMessage += "The beginning odometer miles cannot be a negative number"
                        BeginOdometerTextBox.Focus()
                    Else
                        errorMessage += ", the beginning odometer miles cannot be a negative number"
                    End If
                    valid = False
                    BeginOdometerTextBox.Text = ""
                    BeginOdometerTextBox.BackColor = Color.LightYellow
                End If
            Catch ex As Exception
                'Input was not a number
                If valid Then
                    errorMessage += "The beginning odometer miles must be a number"
                    BeginOdometerTextBox.Focus()
                Else
                    errorMessage += ", the beginning odometer miles must be a number"
                End If
                valid = False
                BeginOdometerTextBox.Text = ""
                BeginOdometerTextBox.BackColor = Color.LightYellow
            End Try
        End If

        'Validates that the ending odometer reading was entered and is a number
        'Also validates that the end odometer miles is greater than the beginning odometer miles
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 
        'If end miles > begin miles then both are cleared and highlighted
        If EndOdometerTextBox.Text = "" Then
            If valid Then
                errorMessage += "Please enter the end odometer miles"
                EndOdometerTextBox.Focus()
            Else
                errorMessage += ", please enter the end odometer miles"
            End If
            valid = False
            EndOdometerTextBox.BackColor = Color.LightYellow
        Else
            'Something was entered and now needs to validated
            Try
                validateNumber = CInt(EndOdometerTextBox.Text)
                endingMiles = CInt(EndOdometerTextBox.Text)
                'end miles is a number and is now checked to ensure that it is greater than begin miles
                If beginningMiles > endingMiles Then
                    If valid Then
                        errorMessage += "The ending miles must be greater than the beginning miles"
                        EndOdometerTextBox.Focus()
                    Else
                        errorMessage += ", the ending miles must be greater than the beginning miles"
                    End If
                    valid = False
                    EndOdometerTextBox.Text = ""
                    EndOdometerTextBox.BackColor = Color.LightYellow
                    BeginOdometerTextBox.Text = ""
                    BeginOdometerTextBox.BackColor = Color.LightYellow
                Else
                    EndOdometerTextBox.BackColor = Color.White
                End If
            Catch ex As Exception
                'entered information was not a number
                If valid Then
                    errorMessage += "The ending odometer miles must be a number"
                    EndOdometerTextBox.Focus()
                Else
                    errorMessage += ", the ending odometer miles must be a number"
                End If
                valid = False
                EndOdometerTextBox.Text = ""
                EndOdometerTextBox.BackColor = Color.LightYellow
            End Try
        End If

        'Validates that the number of days that the customer rented the car was entered, and is an integer, and at least 1 and no more than 45
        'an input that is invalid is: cleared, highlighted light-yellow, the discrepancy is added to errorMessage,
        'if valid is true, then the focus is set to this input, and valid is then set to false to fail the validation process 
        If DaysTextBox.Text = "" Then
            If valid Then
                errorMessage += "Please enter the number of days rented"
                DaysTextBox.Focus()
            Else
                errorMessage += ", please enter the number of days rented"
            End If
            valid = False
            DaysTextBox.Text = ""
            DaysTextBox.BackColor = Color.LightYellow
        Else
            'Something was input and now needs to be check that it is a number at least 1 and no more than 45
            Try
                validateNumber = CInt(DaysTextBox.Text)
                'input is a number and is checked that it is equal to or greater than 1
                If Math.Round(CInt(DaysTextBox.Text)) >= 1 Then
                    'Days is at least 1
                    If CInt(DaysTextBox.Text) > 45 Then
                        'days is greater than 45 and invalid
                        If valid Then
                            errorMessage += "The number of days rented cannot be greater than 45"
                            DaysTextBox.Focus()
                        Else
                            errorMessage += ", the number of days rented cannot be greater than 45"
                        End If
                        valid = False
                        DaysTextBox.Text = ""
                        DaysTextBox.BackColor = Color.LightYellow
                    Else
                        'days is valid
                        DaysTextBox.BackColor = Color.White
                    End If
                Else
                    'days is less than 1 and invalid
                    If valid Then
                        errorMessage += "The number of days rented must be greater than zero"
                        DaysTextBox.Focus()
                    Else
                        errorMessage += ", the number of days rented must be greater than zero"
                    End If
                    valid = False
                    DaysTextBox.Text = ""
                    DaysTextBox.BackColor = Color.LightYellow
                End If
            Catch ex As Exception
                'days input is not a number
                If valid Then
                errorMessage += "The number of days rented must be a whole number"
                DaysTextBox.Focus()
            Else
                errorMessage += ", the number of days rented must be a whole number"
            End If
            valid = False
            DaysTextBox.Text = ""
            DaysTextBox.BackColor = Color.LightYellow
        End Try
        End If

        'if valid then it does not display a message to the user, does if information is invalid
        If valid = False Then
            MsgBox(errorMessage)
        Else
        End If

        Return valid
    End Function

    'Determines the miles driven and the charges for the customer, taking into account any discounts.
    'Puts information to the output text boxes, and stores the totals of the number of customers helped, the miles driven,
    'and the amount charged
    Sub DetermineCost()
        Dim milesDriven As Double = 0
        Dim additionalMiles As Double = 0
        Dim mileageCharge As Double = 0
        Dim dayCharge As Double = 0
        Dim discount As Double = 1
        Dim discountSavings As Double = 0
        Dim total As Double = 0

        'Determines whether or not the distance driven is in km or mi
        'ensure that it is in miles by converting it if needed
        If KilometersradioButton.Checked = True Then
            'if kilometers are selected, then it multiplies the distance by 0.62 to convert it to miles
            milesDriven = (CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)) * 0.62
        Else
            milesDriven = CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)
        End If

        'Rounds the miles driven, and displays the miles driven in the TotalMilesTextBox
        milesDriven = Math.Round(milesDriven, 2)
        TotalMilesTextBox.Text = $"{milesDriven} mi"

        'Determines if there are any charges for additional miles and charges accordingly
        'miles 201-500 are $0.12/mi, miles 501+ are $0.10/mi
        additionalMiles = milesDriven - 200
        If additionalMiles > 0 Then
            'Customer drove over the 200 complementary miles
            If additionalMiles > 300 Then
                'customer drove at least 500 miles, and is charged at $0.10/mi for anything over 500 miles
                'and is charged for 300 miles at $0.12/mi
                mileageCharge = ((additionalMiles - 300) * 0.1) + (300 * 0.12)
            Else
                'customer drove less than 501 miles and is charged for additional miles past 200 miles at $0.12/mi
                mileageCharge = additionalMiles * 0.12
            End If
        Else
            'customer did not exceed the 200 complementary miles and is not charged extra
        End If

        'mileageCharge is rounded to two decimal points, is displayed into MileageChargeTextBox and added to the total
        mileageCharge = Math.Round(mileageCharge, 2)
        MileageChargeTextBox.Text = $"${mileageCharge}"
        total += mileageCharge

        'the daily fee is $15/day, added to the total and put out the DayChargeTextBox
        dayCharge = CInt(DaysTextBox.Text) * 15
        DayChargeTextBox.Text = $"${dayCharge}"
        total += dayCharge

        'Checks discounts that need to be applied
        'AAA members get 5% discount, senior citizens get 3% discount
        'One or both may be applied at a time
        If AAAcheckbox.Checked = True Then
            'Customer is a AAA member
            If Seniorcheckbox.Checked = True Then
                'Customer is also a senior citizen
                discount = 0.92
            Else
                'Customer is not a senior citizen
                discount = 0.95
            End If
        Else
            'customer is not a AAA member
            If Seniorcheckbox.Checked = True Then
                'customer is a senior citizen
                discount = 0.97
            Else
                'customer is not a senior citizen or a AAA member
            End If
        End If

        'discount savings is calculated by subtracting the discounted total from the original total, is rounded to 2 decimal places
        'and output to TotalDiscountTextBox
        discountSavings = Math.Round(total - (total * discount), 2)
        TotalDiscountTextBox.Text = $"${discountSavings}"

        'New discounted total is calculated, rounded to 2 decimal places, and output to TotalChargeTextBox
        'If no discounts were applied, discount = 1, and the total will remain the same
        total = Math.Round((total * discount), 2)
        TotalChargeTextBox.Text = $"${total}"

        'Stores the summary information, adding 1 to store customers, adding the miles driven, and the total charged
        StoreCustomers(1)
        StoreMiles(milesDriven)
        StoreCharges(total)
    End Sub

    'Stores the number of customers helped
    Function StoreCustomers(newCustomer As Double) As Double
        Static storedCustomers As Double

        If newCustomer = -1 Then
            'If newCustomer = -1 then it is being called, and does not add 1 to the amount of customers helped
        Else
            'A new customer was helped and 1 is added to the total customers helped
            storedCustomers += 1
        End If

        Return storedCustomers
    End Function

    'Stores the number of miles traveled by customers
    Function StoreMiles(newMilesDriven As Double) As Double
        Static storedMiles As Double

        If newMilesDriven = -1 Then
            'If newMilesDriven = -1 then it is being called, and nothing is being added to the total
        Else
            'A new customer was helped, and the customers miles driven is added to the stored miles
            storedMiles += newMilesDriven
        End If

        Return storedMiles
    End Function

    'Stores the amount charged from customers
    Function StoreCharges(newCharges As Double) As Double
        Static storedCharges As Double

        If newCharges = -1 Then
            'If newCharges = -1 then it is being called, and nothing is being added to the total
        Else
            'A new customer was helped, and the customers total is added to the storedCharges
            storedCharges += newCharges
        End If

        Return storedCharges
    End Function

    'A detailed summary is displayed, showing the total customers helped, miles driven by customers, and amount charged
    Sub Summary()
        'Calls all of the stored information
        Dim totalCustomers As Double = StoreCustomers(-1)
        Dim totalMilesDriven As Double = StoreMiles(-1)
        Dim totalCharges As Double = StoreCharges(-1)

        'Compiles a message box that displays the information
        MessageBox.Show(($"Total Customers:        {totalCustomers}" & vbCrLf _
                       & $"Total Miles Driven:     {totalMilesDriven} mi" & vbCrLf _
                       & $"Total Charges:            ${totalCharges}"), "Detailed Summary")

    End Sub

    'When exit button is pressed, it displays a pop-up box that asks the user if they would like to close the program
    Function DoYouWantToExit() As Boolean
        Dim leave As Boolean = False
        Dim userInput As Integer = 7

        'Displays the message box, with a yes and no button, yes returns a 6
        userInput = MsgBox("Are you sure that you would like to close this program?", vbYesNo, "Exit?")

        'If yes was pressed, then userInput = 6
        If userInput = 6 Then
            'yes was pressed, and leave = true
            leave = True
        Else
            'no was pressed
        End If

        'if yes was pressed then returns a true, else it returns a false
        Return leave
    End Function

    'Event Handlers below this point

    'Input for customer name
    Private Sub NameTextBox_TextChanged(sender As Object, e As EventArgs) Handles NameTextBox.TextChanged

    End Sub

    'Input for customer address
    Private Sub AddressTextBox_TextChanged(sender As Object, e As EventArgs) Handles AddressTextBox.TextChanged

    End Sub

    'Input for customer city
    Private Sub CityTextBox_TextChanged(sender As Object, e As EventArgs) Handles CityTextBox.TextChanged

    End Sub

    'Input for customer state
    Private Sub StateTextBox_TextChanged(sender As Object, e As EventArgs) Handles StateTextBox.TextChanged

    End Sub

    'Input for customer zip code
    Private Sub ZipCodeTextBox_TextChanged(sender As Object, e As EventArgs) Handles ZipCodeTextBox.TextChanged

    End Sub

    'Input for beginning odometer miles
    Private Sub BeginOdometerTextBox_TextChanged(sender As Object, e As EventArgs) Handles BeginOdometerTextBox.TextChanged

    End Sub

    'Input for end odometer miles
    Private Sub EndOdometerTextBox_TextChanged(sender As Object, e As EventArgs) Handles EndOdometerTextBox.TextChanged

    End Sub

    'Input for the number of days that the customer rented a vehicle
    Private Sub DaysTextBox_TextChanged(sender As Object, e As EventArgs) Handles DaysTextBox.TextChanged

    End Sub

    'File button on the top menu strip
    Private Sub FileToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem1.Click

    End Sub

    'Calculate button located in the file drop down on the top menu strip
    'Validates information, and if validated, then the summary buttons are enabled, and the costs are determined,
    'then put to the output
    Private Sub CalculateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateToolStripMenuItem.Click
        If ValidateInputs() Then
            SummaryButton.Enabled = True
            SummaryToolStripMenuItem1.Enabled = True
            DetermineCost()
        Else
        End If
    End Sub

    'Clear button located in the file drop down on the top menu strip
    'Clears the customer information and the output, resets check boxes to default
    Private Sub ClearToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem1.Click
        Defaults()
    End Sub

    'Summary button located in the file drop down on the top menu strip
    'Displays the detailed summary of the totals of customers helped, the miles driven, and the charges
    'Disabled until valid customer information is input
    Private Sub SummaryToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SummaryToolStripMenuItem1.Click
        Defaults()
        Summary()
    End Sub

    'Exit button located in the file drop down on the top menu strip
    'Asks the user if they want to leave, directing them to the DoYouWantToExit() function
    'A True return closes the form, a false cancels the action
    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        If DoYouWantToExit() Then
            Me.Close()
        Else
        End If
    End Sub

    'Label for NameTextBox
    Private Sub NameLabel_Click(sender As Object, e As EventArgs) Handles NameLabel.Click

    End Sub

    'Label for AddressTextBox
    Private Sub AddressLabel_Click(sender As Object, e As EventArgs) Handles AddressLabel.Click

    End Sub

    'Form initializing sub
    'Sets the default settings, with the summary buttons disabled
    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Defaults()
        SummaryButton.Enabled = False
        SummaryToolStripMenuItem1.Enabled = False
    End Sub

    'Label for CityTextBox
    Private Sub CityLabel_Click(sender As Object, e As EventArgs) Handles CityLabel.Click

    End Sub

    'Label for StateTextBox
    Private Sub State4_Click(sender As Object, e As EventArgs) Handles State4.Click

    End Sub

    'Label for ZipCodeTextBox
    Private Sub ZipCodeLabel_Click(sender As Object, e As EventArgs) Handles ZipCodeLabel.Click

    End Sub

    'Label for BeginOdometerTextBox
    Private Sub BeginOdometerLabel_Click(sender As Object, e As EventArgs) Handles BeginOdometerLabel.Click

    End Sub

    'Label for EndOdometerTextBox
    Private Sub EndOdometerLabel_Click(sender As Object, e As EventArgs) Handles EndOdometerLabel.Click

    End Sub

    'Label for DaysTextBox
    Private Sub DaysLabel_Click(sender As Object, e As EventArgs) Handles DaysLabel.Click

    End Sub

    'GroupBox that houses the customer information inputs
    Private Sub OdometerGroupbox_Enter(sender As Object, e As EventArgs) Handles OdometerGroupbox.Enter

    End Sub

    'Radio button that signals that the odometer distance inputs are in miles and not kilometers
    Private Sub MilesradioButton_CheckedChanged(sender As Object, e As EventArgs) Handles MilesradioButton.CheckedChanged

    End Sub

    'Radio button that signals that the odometer distance inputs are in kilometers and not miles
    Private Sub KilometersradioButton_CheckedChanged(sender As Object, e As EventArgs) Handles KilometersradioButton.CheckedChanged

    End Sub

    'Label for the TotalMilesTextBox
    Private Sub MilesDrivenLabel_Click(sender As Object, e As EventArgs) Handles MilesDrivenLabel.Click

    End Sub

    'Label for the MileChargeTextBox output
    Private Sub MileChargeLabel_Click(sender As Object, e As EventArgs) Handles MileChargeLabel.Click

    End Sub

    'Label for the DayChargeTextBox output
    Private Sub DayChargeLabel_Click(sender As Object, e As EventArgs) Handles DayChargeLabel.Click

    End Sub

    'Label for the TotalDiscountTextBox output
    Private Sub DiscountLabel_Click(sender As Object, e As EventArgs) Handles DiscountLabel.Click

    End Sub

    'Label for the TotalChargeTextBox output
    Private Sub YouOweLabel_Click(sender As Object, e As EventArgs) Handles YouOweLabel.Click

    End Sub

    'TotalMilesTextBox output to display miles driven
    Private Sub TotalMilesTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalMilesTextBox.TextChanged

    End Sub

    'MileageChargeTextBox output to display fees for additional miles
    Private Sub MileageChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles MileageChargeTextBox.TextChanged

    End Sub

    'DaysChargeTextBox output to display the fee for number of days rented
    Private Sub DayChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles DayChargeTextBox.TextChanged

    End Sub

    'TotalDiscountTextBox output to display the amount of savings from applied discounts
    Private Sub TotalDiscountTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalDiscountTextBox.TextChanged

    End Sub

    'TotalChargeTextBox output to display the total being charged to the customer
    Private Sub TotalChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalChargeTextBox.TextChanged

    End Sub

    'DiscountGroupBox that houses the discount check boxes
    Private Sub DiscountGroupbox_Enter(sender As Object, e As EventArgs) Handles DiscountGroupbox.Enter

    End Sub

    'Check box that applies the AAA member discount of 5%
    Private Sub AAAcheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles AAAcheckbox.CheckedChanged

    End Sub

    'Check box that applies the senior citizen discount of 3%
    Private Sub Seniorcheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles Seniorcheckbox.CheckedChanged

    End Sub

    'Validates information, and if validated, then the summary buttons are enabled, and the costs are determined,
    'then put to the output
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If ValidateInputs() Then
            SummaryButton.Enabled = True
            SummaryToolStripMenuItem1.Enabled = True
            DetermineCost()
        Else
        End If
    End Sub

    'Clears the customer information and the output, resets check boxes to default
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Defaults()
    End Sub

    'Displays the detailed summary of the totals of customers helped, the miles driven, and the charges
    'Disabled until valid customer information is input
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Defaults()
        Summary()
    End Sub

    'Asks the user if they want to leave, directing them to the DoYouWantToExit() function
    'A True return closes the form, a false cancels the action
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If DoYouWantToExit() Then
            Me.Close()
        Else
        End If
    End Sub
End Class

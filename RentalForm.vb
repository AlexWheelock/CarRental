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
    End Sub


    Function ValidateInputs() As Boolean
        Dim valid As Boolean = True
        Dim validateString As Integer
        Dim validateAddress() As String = Split(AddressTextBox.Text, " ")
        Dim validateNumber As Integer
        Dim beginningMiles As Integer
        Dim endingMiles As Integer
        Dim errorMessage As String = ("The following information is incorrect:" & vbCrLf _
            & vbCrLf)

        'Validates that the name is not a number
        Try
            validateString = CInt(NameTextBox.Text)
            valid = False
            NameTextBox.Focus()
            NameTextBox.BackColor = Color.LightYellow
            NameTextBox.Text = ""
            errorMessage += "Name cannot contain a number"
        Catch ex As Exception
            If NameTextBox.Text = "" Then
                valid = False
                NameTextBox.Focus()
                NameTextBox.BackColor = Color.LightYellow
                errorMessage += "Please enter a name"
            Else
                NameTextBox.BackColor = Color.White
            End If
        End Try


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
            Try
                validateNumber = CInt(validateAddress(0))
                AddressTextBox.BackColor = Color.White
            Catch ex As Exception
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
        End If

        'Validates that the city is not a number
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
            Try
                validateString = CInt(CityTextBox.Text)
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
                CityTextBox.BackColor = Color.White
            End Try
        End If

        'Validates that the state is not a number
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
                StateTextBox.BackColor = Color.White
            End Try
        End If

        'Validate that the Zip is a number
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
            Try
                validateNumber = CInt(ZipCodeTextBox.Text)
                ZipCodeTextBox.BackColor = Color.White
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
                End If
            Catch ex As Exception
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

        'Validates that the beginning odometer reading is a number
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
            Try
                validateNumber = CInt(BeginOdometerTextBox.Text)
                beginningMiles = CInt(BeginOdometerTextBox.Text)
                BeginOdometerTextBox.BackColor = Color.White
            Catch ex As Exception
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

        'Validates that the ending odometer reading is a number
        'Also validates that the end odometer miles is greater than the beginning odometer miles
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
            Try
                validateNumber = CInt(EndOdometerTextBox.Text)
                endingMiles = CInt(EndOdometerTextBox.Text)
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

        'Validates that the number of days that the customer rented the car was
        'an integer, and at least 1
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
            Try
            validateNumber = CInt(DaysTextBox.Text)
            If CInt(DaysTextBox.Text) > 0 Then
                DaysTextBox.BackColor = Color.White
            ElseIf CInt(DaysTextBox.Text) > 45 Then
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

        If valid = False Then
            MsgBox(errorMessage)
        Else
        End If

        Return valid
    End Function

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
            milesDriven = (CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)) * 0.62
        Else
            milesDriven = CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)
        End If

        TotalMilesTextBox.Text = $"{milesDriven} mi"

        'Determines if there are any charges for additional miles and charges accordingly
        'miles 201-500 are $0.12/mi, miles 501+ are $0.10/mi
        additionalMiles = milesDriven - 200
        If additionalMiles > 0 Then
            If additionalMiles > 300 Then
                mileageCharge = ((additionalMiles - 300) * 0.1) + (300 * 0.12)
            Else
            End If
        Else
            mileageCharge = additionalMiles * 0.12
        End If

        MileageChargeTextBox.Text = $"${mileageCharge}"
        total += mileageCharge

        'the daily fee is $15/day, added to the total and put out the the DayChargeTextBox
        dayCharge = CInt(DaysTextBox.Text) * 15
        DayChargeTextBox.Text = $"${dayCharge}"
        total += dayCharge

        'Checks discounts that need to be applied
        'AAA members get 5% discount, senior citizens get 3% discount
        'One or both may be applied at a time
        If AAAcheckbox.Checked = True Then
            If Seniorcheckbox.Checked = True Then
                discount = 0.92
            Else
                discount = 0.95
            End If
        Else
            If Seniorcheckbox.Checked = True Then
                discount = 0.97
            Else
            End If
        End If

        discountSavings = total - (total * discount)
        TotalDiscountTextBox.Text = $"${discountSavings}"

        total = (total * discount)
        TotalChargeTextBox.Text = $"${total}"

        StoreCustomers(1)
        StoreMiles(milesDriven)
        StoreCharges(total)
    End Sub

    Function StoreCustomers(newCustomer As Double) As Double
        Static storedCustomers As Double

        If newCustomer = -1 Then
        Else
            storedCustomers += 1
        End If

        Return storedCustomers
    End Function

    Function StoreMiles(newMilesDriven As Double) As Double
        Static storedMiles As Double

        If newMilesDriven = -1 Then
        Else
            storedMiles += newMilesDriven
        End If

        Return storedMiles
    End Function

    Function StoreCharges(newCharges As Double) As Double
        Static storedCharges As Double

        If newCharges = -1 Then
        Else
            storedCharges += newCharges
        End If

        Return storedCharges
    End Function

    Sub Summary()
        Dim totalCustomers As Double = StoreCustomers(-1)
        Dim totalMilesDriven As Double = StoreMiles(-1)
        Dim totalCharges As Double = StoreCharges(-1)

        MessageBox.Show(($"Total Customers:          {totalCustomers}" & vbCrLf _
               & $"Total Miles Driven:          {totalMilesDriven} mi" & vbCrLf _
               & $"Total Charges:         ${totalCharges}"), "Detailed Summary")

    End Sub

    'Event Handlers below this point

    Private Sub NameTextBox_TextChanged(sender As Object, e As EventArgs) Handles NameTextBox.TextChanged

    End Sub

    Private Sub AddressTextBox_TextChanged(sender As Object, e As EventArgs) Handles AddressTextBox.TextChanged

    End Sub

    Private Sub CityTextBox_TextChanged(sender As Object, e As EventArgs) Handles CityTextBox.TextChanged

    End Sub

    Private Sub StateTextBox_TextChanged(sender As Object, e As EventArgs) Handles StateTextBox.TextChanged

    End Sub

    Private Sub ZipCodeTextBox_TextChanged(sender As Object, e As EventArgs) Handles ZipCodeTextBox.TextChanged

    End Sub

    Private Sub BeginOdometerTextBox_TextChanged(sender As Object, e As EventArgs) Handles BeginOdometerTextBox.TextChanged

    End Sub

    Private Sub EndOdometerTextBox_TextChanged(sender As Object, e As EventArgs) Handles EndOdometerTextBox.TextChanged

    End Sub

    Private Sub DaysTextBox_TextChanged(sender As Object, e As EventArgs) Handles DaysTextBox.TextChanged

    End Sub

    Private Sub FileToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem1.Click

    End Sub

    Private Sub CalculateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateToolStripMenuItem.Click

    End Sub

    Private Sub ClearToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem1.Click

    End Sub

    Private Sub SummaryToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SummaryToolStripMenuItem1.Click

    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click

    End Sub

    Private Sub NameLabel_Click(sender As Object, e As EventArgs) Handles NameLabel.Click

    End Sub

    Private Sub AddressLabel_Click(sender As Object, e As EventArgs) Handles AddressLabel.Click

    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Defaults()
        SummaryButton.Enabled = False
    End Sub

    Private Sub CityLabel_Click(sender As Object, e As EventArgs) Handles CityLabel.Click

    End Sub

    Private Sub State4_Click(sender As Object, e As EventArgs) Handles State4.Click

    End Sub

    Private Sub ZipCodeLabel_Click(sender As Object, e As EventArgs) Handles ZipCodeLabel.Click

    End Sub

    Private Sub BeginOdometerLabel_Click(sender As Object, e As EventArgs) Handles BeginOdometerLabel.Click

    End Sub

    Private Sub EndOdometerLabel_Click(sender As Object, e As EventArgs) Handles EndOdometerLabel.Click

    End Sub

    Private Sub DaysLabel_Click(sender As Object, e As EventArgs) Handles DaysLabel.Click

    End Sub

    Private Sub OdometerGroupbox_Enter(sender As Object, e As EventArgs) Handles OdometerGroupbox.Enter

    End Sub

    Private Sub MilesradioButton_CheckedChanged(sender As Object, e As EventArgs) Handles MilesradioButton.CheckedChanged

    End Sub

    Private Sub KilometersradioButton_CheckedChanged(sender As Object, e As EventArgs) Handles KilometersradioButton.CheckedChanged

    End Sub

    Private Sub MilesDrivenLabel_Click(sender As Object, e As EventArgs) Handles MilesDrivenLabel.Click

    End Sub

    Private Sub MileChargeLabel_Click(sender As Object, e As EventArgs) Handles MileChargeLabel.Click

    End Sub

    Private Sub DayChargeLabel_Click(sender As Object, e As EventArgs) Handles DayChargeLabel.Click

    End Sub

    Private Sub DiscountLabel_Click(sender As Object, e As EventArgs) Handles DiscountLabel.Click

    End Sub

    Private Sub YouOweLabel_Click(sender As Object, e As EventArgs) Handles YouOweLabel.Click

    End Sub

    Private Sub TotalMilesTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalMilesTextBox.TextChanged

    End Sub

    Private Sub MileageChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles MileageChargeTextBox.TextChanged

    End Sub

    Private Sub DayChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles DayChargeTextBox.TextChanged

    End Sub

    Private Sub TotalDiscountTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalDiscountTextBox.TextChanged

    End Sub

    Private Sub TotalChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalChargeTextBox.TextChanged

    End Sub

    Private Sub DiscountGroupbox_Enter(sender As Object, e As EventArgs) Handles DiscountGroupbox.Enter

    End Sub

    Private Sub AAAcheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles AAAcheckbox.CheckedChanged

    End Sub

    Private Sub Seniorcheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles Seniorcheckbox.CheckedChanged

    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If ValidateInputs() Then
            SummaryButton.Enabled = True
            DetermineCost()
        Else
        End If
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Defaults()
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Defaults()
        Summary()
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        ExitForm.ShowDialog()
    End Sub
End Class

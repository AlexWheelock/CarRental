'Alex Wheelock
'RCET 0625
'Car Rental
'Spring 2024
'https://github.com/AlexWheelock/CarRental

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm

    Sub ValidateInputs()
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
            NameTextBox.BackColor = Color.White
        End Try

        'Validates that the address has a home number at the beginning
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

        'Validates that the address has a street name
        Try
            validateString = CInt(validateAddress(1))
            If valid Then
                errorMessage += "The address must contain a street name"
                AddressTextBox.Focus()
            Else
                errorMessage += ", the address must contain a street name"
            End If
            valid = False
            AddressTextBox.Text = ""
            AddressTextBox.BackColor = Color.LightYellow
        Catch ex As Exception
            AddressTextBox.BackColor = Color.White
        End Try

        'Validates that the city is not a number
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

        'Validates that the state is not a number
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

        'Validate that the Zip is a number
        Try
            validateNumber = CInt(ZipCodeTextBox.Text)
            ZipCodeTextBox.BackColor = Color.White
        Catch ex As Exception
            If valid Then
                errorMessage += "The zip code must be a number"
                ZipCodeTextBox.Focus()
            Else
                errorMessage += ", the zip code must number"
            End If
            valid = False
            ZipCodeTextBox.Text = ""
            ZipCodeTextBox.BackColor = Color.LightYellow
        End Try

        'Validates that the beginning odometer reading is a number
        Try
            validateNumber = CInt(BeginOdometerTextBox.Text)
            beginningMiles = CInt(BeginOdometerTextBox.Text)
            BeginOdometerTextBox.BackColor = Color.White
        Catch ex As Exception
            If valid Then
                errorMessage += "The beginning odometer miles must be a number"
                BeginOdometerTextBox.Focus()
            Else
                errorMessage += ", the beginning odometer miles must number"
            End If
            valid = False
            BeginOdometerTextBox.Text = ""
            BeginOdometerTextBox.BackColor = Color.LightYellow
        End Try

        'Validates that the ending odometer reading is a number
        'Also validates that the end odometer miles is greater than the beginning odometer miles
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
                errorMessage += ", the ending odometer miles must number"
            End If
            valid = False
            EndOdometerTextBox.Text = ""
            EndOdometerTextBox.BackColor = Color.LightYellow
        End Try

    End Sub

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

    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click

    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click

    End Sub
End Class

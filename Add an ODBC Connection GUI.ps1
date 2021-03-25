#region Initial Declarations

#-- Initial WPF Declaration
Add-Type -AssemblyName PresentationCore, PresentationFramework

#endregion

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="800" Height="400">
<Grid Margin="1,1,-1,-1">
    <Grid.ColumnDefinitions>
        <ColumnDefinition/>
        <ColumnDefinition/>
        <ColumnDefinition/>
    </Grid.ColumnDefinitions>
    
    <Grid.RowDefinitions>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
        <RowDefinition/>
    </Grid.RowDefinitions>
    
 
<Button Content="Create Connections" Grid.Row="13" Grid.Column="1" Grid.RowSpan="2" Name="ButtonCreateConnections" Background="DarkCyan" FontWeight="Bold" Foreground="White" FontSize="16"/>

<Label Content=" Add an ODBC Connection" Grid.RowSpan="2" Grid.ColumnSpan="2" FontSize="20"/>

<CheckBox Content="ODBC Connection 1" Grid.Row="3" Grid.Column="1" Name="CheckBox1" FontSize="15"/>

<CheckBox Content="ODBC Connection 2" Grid.Row="5" Grid.Column="1" Name="CheckBox2" FontSize="15"/>

<CheckBox Content="ODBC Connection 3" Grid.Row="7" Grid.Column="1" Name="CheckBox3" FontSize="15"/>

<CheckBox Content="ODBC Connection 4" Grid.Row="9" Grid.Column="1" Name="CheckBox4" FontSize="15"/>

<CheckBox Content="ODBC Connection 5" Grid.Row="11" Grid.Column="1" Name="CheckBox5" FontSize="15"/>

</Grid>
</Window>
"@


#region Xaml WPF config

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

#endregion

#region Script Logic

#-- Button click event

$ButtonCreateConnections.Add_Click({ButtonCreateConnections_Click $this $_})

function ButtonCreateConnections_Click {

        $test1 = ""
        $test2 = ""
        $test3 = ""
        $test4 = ""
        $test5 = ""

        #-- Each if statement only applies if that checkbox is checked when the button is clicked
        #-- If an ODBC connection with that connection name already exists, do nothing and display error. Otherwise create the ODBC connection with that name

        if($CheckBox1.isChecked) 
        {
        $nameCheck1 = Get-OdbcDsn -Name "ODBC Connection 1"
        if($nameCheck1 -eq $null)
            {
            Add-OdbcDsn -Name "ODBC Connection 1" -DriverName "SQL Server" -DsnType "user" -SetPropertyValue @("Server=SERVERNAME")
            $test1 = "ODBC Connection 1"
            }
            else
            {
            [System.Windows.MessageBox]::Show("An ODBC Connection named ODBC Connection 1 Already Exists")
            }
        }

        if($CheckBox2.isChecked) 
        {
        $nameCheck2 = Get-OdbcDsn -Name "ODBC Connection 2"
        if($nameCheck2 -eq $null)
            {
            Add-OdbcDsn -Name "ODBC Connection 2" -DriverName "SQL Server" -DsnType "user" -SetPropertyValue @("Server=SERVERNAME")
            $test2 = "ODBC Connection 2"
            }
            else
            {
            [System.Windows.MessageBox]::Show("An ODBC Connection named ODBC Connection 2 Already Exists")
            }
        }

        if($CheckBox3.isChecked) 
        {
        $nameCheck3 = Get-OdbcDsn -Name "ODBC Connection 3"
        if($nameCheck3 -eq $null)
            {
            Add-OdbcDsn -Name "ODBC Connection 3" -DriverName "SQL Server" -DsnType "user" -SetPropertyValue @("Server=SERVERNAME")
            $test3 = "ODBC Connection 3"
            }
            else
            {
            [System.Windows.MessageBox]::Show("An ODBC Connection named ODBC Connection 3 Already Exists")
            }
        }

        if($CheckBox4.isChecked) 
        {
        $nameCheck4 = Get-OdbcDsn -Name "ODBC Connection 4"
        if($nameCheck4 -eq $null)
            {
            Add-OdbcDsn -Name "ODBC Connection 4" -DriverName "SQL Server" -DsnType "user" -SetPropertyValue @("Server=SERVERNAME")
            $nameCheck4 = "ODBC Connection 4"
            }
            else
            {
            [System.Windows.MessageBox]::Show("An ODBC Connection named ODBC Connection 4 Already Exists")
            }
        }

        if($CheckBox5.isChecked) 
        {
        $nameCheck5 = Get-OdbcDsn -Name "ODBC Connection 5"
        if($nameCheck5 -eq $null)
            {
            Add-OdbcDsn -Name "ODBC Connection 5" -DriverName "SQL Server" -DsnType "user" -SetPropertyValue @("Server=SERVERNAME")
            $test5 = "ODBC Connection 5"
            }
            else
            {
            [System.Windows.MessageBox]::Show("An ODBC Connection named ODBC Connection 5 Already Exists")
            }
        }

        #-- if no connections have been created, display first message. 
        #-- if any connections have been created, display second message which includes only the name of the connection(s) that have been created.

        if($test1 -eq "" -and $test2 -eq "" -and $test3 -eq "" -and $test4 -eq "" -and $test5 -eq "")
        {
        [System.Windows.MessageBox]::Show("No ODBC Connections were created")
        }
        else
        {
        [System.Windows.MessageBox]::Show("The ODBC Connections:`n $test1   $test2   $test3   $test4   $test5 `nHave sucessfully been created")
        }
        

}

#endregion

$Window.ShowDialog()
# Load presentation framework assembly
Add-Type -AssemblyName PresentationFramework

# Define XAML layout
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Folder Processor" Height="450" Width="800">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <TextBox Name="SearchTextBox" Width="250" Height="30" Margin="5" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center"/>
        <Button Name="BrowseButton" Content="Browse" Width="100" Height="30" Margin="5" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Right" VerticalAlignment="Center"/>
        <ListBox Name="FolderListBox" Grid.Row="1" Grid.Column="0" Margin="5" SelectionMode="Multiple"/>
        <Button Name="ProcessButton" Content="Process Folders" Width="150" Height="30" Margin="5" HorizontalAlignment="Right" VerticalAlignment="Center" Grid.Row="2" Grid.Column="0"/>
    </Grid>
</Window>
"@

function move-binders {
    param(
        [string]$location,
        [string]$destination
    )
    $objShell = New-Object -ComObject 'Shell.Application'
    $destinationFolder = $objShell.NameSpace($destination)
    $destinationFolder.CopyHere($location)
}

# Load XAML into PowerShell
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Access XAML elements
$searchTextBox = $window.FindName("SearchTextBox")
$browseButton = $window.FindName("BrowseButton")
$folderListBox = $window.FindName("FolderListBox")
$processButton = $window.FindName("ProcessButton")

# Initialize a variable to store the original folder list
$script:folderList = @()

# Browse Button Click event
$browseButton.Add_Click({
    $folderListBox.items.Clear()
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder"
    $result = $folderBrowser.ShowDialog()

    if ($result -eq "OK") {
        $folderPath = $folderBrowser.SelectedPath
        Get-ChildItem -Path $folderPath -Directory -depth 1 | % {$script:folderList += $_}
        $folderListBox.Items.Clear()
        foreach ($folder in $folderList) {
            $folderListBox.Items.Add($folder.Name)
        }
    }
})

# Search TextBox TextChanged event
$searchTextBox.Add_TextChanged({
    $searchText = $searchTextBox.Text
    
    if ($searchText.length -ge 4) {
        $folderListBox.items.clear()
        $script:folderList | % {if($_ -match $searchText){$folderListBox.items.add($_)}}
    } else {
        foreach ($folder in $folderList) {
        $folderListBox.Items.Add($folder.name)
        }
    }

})

# Process Button Click event
$processButton.Add_Click({


  # Define the location list of directories for which the files will be sorted
  $sortingOrder = @(
    #removed for data privacy
  )

  # Get the selected items from the listbox
  $selectedItems = $binderlist.SelectedItems

  # Loop through each selected item and move it to the appropriate directory
  foreach ($item in $selectedItems) {
    # Extract the year and first letter from the item name using regex
    $itemName = ($item -split "\\")[-1]
    $year = [regex]::Match($itemName, "(\d{4})-\d{2}-\d{2}")
    $year = $year.groups[1].value
    $firstLetter = $itemName.substring(0,1)

    $matchIndex = -1
    foreach ($dir in $sortingOrder) {
      if ($dir -match "$year ([0-9A-Za-z])-([A-Za-z])?") {
        $rangeStart = $Matches[1]
        $rangeEnd = $Matches[2]
        if ($firstLetter -ge $rangeStart) {
          if ($firstLetter -le $rangeEnd) {
            $matchIndex = $sortingOrder.IndexOf($dir)
            break
          }
        }
      }
    }

    move-binders -location $item  -destination $sortingOrder[$matchIndex]
    }
})

# Show the window
$window.ShowDialog() | Out-Null

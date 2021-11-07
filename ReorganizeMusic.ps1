<#
  This script takes an input folder of mp3, flac, and m4a files and creates an output folder within it called 'Music' which
  contains copies of all files found but whose naming conventions and hierarchy match those required by Plex Media Server.

  You can read more about the required hierarchy here: https://support.plex.tv/articles/200265296-adding-music-media-from-folders/

  Specifically, this script creates this hierarchy:

    Content should have each artist in their own directory, with each album as a separate subdirectory within it.

      Music/ArtistName/AlbumName/TrackNumber - TrackName.ext
#>
param(
  [string] $folder
)

<#
  Sanitizes a string to not contain any invalid file name characters
#>
function Sanitize {
  param($inputString)

  if($inputString -ne $null)
  {
    $inputString.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
  }
}

<#
  Takes a single full path of a track and uses the metadata on the file to create a new hierarchy and name for
  the track under the Music folder.
#>
function ProcessSingleTrack {

  param($track)

  $folder = Split-Path $track
  $file = Split-Path $track -leaf

  $parts = (Split-Path -Path $track -leaf).Split(".")
  $extension = "." + $parts[$parts.Length - 1];

  $shellfolder = $shell.Namespace($folder)
  $shellfile = $shellfolder.ParseName($file)

  $newPath = ""
  $artistName = ""
  $albumName = ""
  $trackNumber = ""
  $trackName = ""

  # Note: Could be extended to allow preferences to be set by user
  $artistName = Sanitize($shellfile.ExtendedProperty("System.Music.AlbumArtist"))
  if($artistName -eq $null)
  {
    $artistName = $shellfile.ExtendedProperty("System.Music.Artist")
  }

  $albumName = Sanitize($shellfile.ExtendedProperty("System.Music.AlbumTitle"))
  $trackNumber = $shellfile.ExtendedProperty("System.Music.TrackNumber")

  $trackName = Sanitize($shellfile.ExtendedProperty("System.Title"))
  if($trackName -eq $null)
  {
    $trackName = Sanitize($shellfile.ExtendedProperty("System.ItemName"))
  }

  if($artistName -eq $null)
  {
    $artistName = "Unknown Artist"
  }

  if($albumName -eq $null)
  {
    $albumName = "Unknown Album"
  }

  if($trackNumber -eq $null)
  {
    $trackNumber = ""
  }
  else
  {
    $trackNumber = ("" + $trackNumber + " - ")    
  }

  if($trackName -eq $null)
  {
    $trackNumber = "Unknown Track"
  }

  if($trackName.EndsWith($extension) -ne $true)
  {
    $trackName += $extension
  }

  # Note: Could add a "WhatIf" mode here
  $newFile = $folder + "\Music\$artistName\$albumName\$trackNumber$trackName"
  New-Item -ItemType File -Path $newFile -Force | Out-Null
  Write-Host $track "To" $newFile
  Copy-Item -LiteralPath $track $newFile
}

# Main
$shell = New-Object -ComObject Shell.Application
gci $folder | % {

  # Note: Could be extended for other media files, but wasn't necessary in my use case
  if($_.extension -eq ".mp3" -or
     $_.extension -eq ".flac" -or
     $_.extension -eq ".m4a")
  {
    ProcessSingleTrack($_.FullName)
  }
}
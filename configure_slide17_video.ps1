param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath
)

$ErrorActionPreference = "Stop"
$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open(
        (Resolve-Path -LiteralPath $PresentationPath).Path,
        $false,
        $false,
        -1
    )

    $video = $presentation.Slides.Item(17).Shapes.Item("Slide17_Video")
    $playSettings = $video.AnimationSettings.PlaySettings
    # These legacy media properties are not exposed consistently by every
    # PowerPoint build. The timeline trigger below is the authoritative
    # autoplay setting; retain loop/rewind when the properties are available.
    try { $playSettings.PlayOnEntry = -1 } catch { }
    try { $playSettings.LoopUntilStopped = -1 } catch { }
    try { $playSettings.RewindMovie = -1 } catch { }

    # PowerPoint can retain an on-click trigger for media inserted by
    # python-pptx even when PlayOnEntry is enabled. Move the media effect into
    # the automatic timeline so it starts as soon as the slide is entered.
    $sequence = $presentation.Slides.Item(17).TimeLine.MainSequence
    for ($index = $sequence.Count; $index -ge 1; $index--) {
        $effect = $sequence.Item($index)
        if ($effect.Shape.Name -eq "Slide17_Video") {
            $effect.Delete()
        }
    }
    # Recreate the media-play effect with an automatic trigger. PowerPoint
    # does not permit changing TriggerType in place for some media effects.
    $autoPlayEffect = $sequence.AddEffect(
        $video,
        83,  # msoAnimEffectMediaPlay
        0,   # msoAnimateLevelNone
        2    # msoAnimTriggerWithPrevious
    )
    $autoPlayEffect.Timing.TriggerDelayTime = 0

    $presentation.Save()
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    while ($presentation.Saved -ne -1 -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 200
    }
    if ($presentation.Saved -ne -1) {
        throw "PowerPoint did not finish saving within 30 seconds."
    }

    Write-Output "Slide 17 video configured for native autoplay and looping."
}
finally {
    if ($null -ne $presentation) {
        $presentation.Close()
    }
    if ($null -ne $powerPoint) {
        $powerPoint.Quit()
    }
}

param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [string]$SeriousMascotPath = (
        Join-Path $PSScriptRoot "assets\Time_Traveler_Spoilers_Red_Portal_Transparent.png"
    )
)

$ErrorActionPreference = "Stop"

$targetSlides = @(1, 4, 5, 10, 11)
$oldBubbleParts = @(
    "SpeechBubble_Background",
    "SpeechBubble_Text",
    "SpeechBubble_Tail"
)
$groupParts = [object[]]@(
    "SpeechBubble_Callout",
    "Mascot",
    "Mascot_Pivot_Frame"
)
$calloutWidths = @{
    1 = 3.55
    4 = 4.36
    5 = 5.76
    10 = 6.12
    11 = 5.00
}
$calloutHeights = @{
    1 = 0.82
    4 = 0.72
    5 = 0.82
    10 = 0.98
    11 = 0.82
}

$pointsPerInch = 72.0
$mascotX = 11.75 * $pointsPerInch
$mascotY = 5.85 * $pointsPerInch
$mascotSize = 1.48 * $pointsPerInch
$mascotOriginX = (11.75 + (1.48 / 2.0)) * $pointsPerInch
$mascotOriginY = (5.85 + (1.48 / 2.0)) * $pointsPerInch

# PowerPoint's native callout pointer extends outside the shape bounds. These
# ratios keep its endpoint fixed above the mascot even when bubble widths differ.
$pointerTargetX = $mascotOriginX
$pointerTargetY = ($mascotY - (0.04 * $pointsPerInch))
$pointerFractionX = 0.92
$pointerExtensionY = 0.35
$animationDuration = 0.25

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open(
        (Resolve-Path -LiteralPath $PresentationPath).Path,
        $false,
        $false,
        $false
    )

    foreach ($slideNumber in $targetSlides) {
        if ($slideNumber -gt $presentation.Slides.Count) {
            continue
        }

        $slide = $presentation.Slides.Item($slideNumber)
        $sequence = $slide.TimeLine.MainSequence

        # Remove the previous click effect before changing or ungrouping shapes.
        for ($effectIndex = $sequence.Count; $effectIndex -ge 1; $effectIndex--) {
            $existingEffect = $sequence.Item($effectIndex)
            if (
                $null -ne $existingEffect.Shape -and
                $existingEffect.Shape.Name -eq "SpeechOverlay_Group"
            ) {
                $existingEffect.Delete()
            }
        }

        try {
            $oldGroup = $slide.Shapes.Item("SpeechOverlay_Group")
            $null = $oldGroup.Ungroup()
        }
        catch {
            # A fresh generator run has not grouped these shapes yet.
        }

        # Remove earlier pivot experiments before rebuilding the group.
        foreach ($pivotShapeName in @("Mascot_Origin", "Mascot_Pivot_Frame")) {
            try {
                $slide.Shapes.Item($pivotShapeName).Delete()
            }
            catch {
                # The shape is absent on a fresh build.
            }
        }

        if ($slideNumber -eq 11) {
            $slide.Shapes.Item("Mascot").Delete()
            $seriousMascot = $slide.Shapes.AddPicture(
                (Resolve-Path -LiteralPath $SeriousMascotPath).Path,
                $false,
                $true,
                $mascotX,
                $mascotY,
                $mascotSize,
                $mascotSize
            )
            $seriousMascot.Name = "Mascot"
        }

        $callout = $null
        try {
            $callout = $slide.Shapes.Item("SpeechBubble_Callout")
        }
        catch {
            $background = $slide.Shapes.Item("SpeechBubble_Background")
            $textShape = $slide.Shapes.Item("SpeechBubble_Text")
            $tail = $slide.Shapes.Item("SpeechBubble_Tail")

            $bubbleText = $textShape.TextFrame.TextRange.Text
            $fontSize = $textShape.TextFrame.TextRange.Font.Size
            if ($fontSize -le 0) {
                $fontSize = 15.5
            }

            $calloutWidth = $calloutWidths[$slideNumber] * $pointsPerInch
            $calloutHeight = $calloutHeights[$slideNumber] * $pointsPerInch
            $calloutLeft = $pointerTargetX - ($pointerFractionX * $calloutWidth)
            $calloutTop = (
                $pointerTargetY -
                ($pointerExtensionY * $calloutHeight) -
                $calloutHeight
            )

            foreach ($shapeName in $oldBubbleParts) {
                $slide.Shapes.Item($shapeName).Delete()
            }

            # msoShapeRoundedRectangularCallout = 106.
            $callout = $slide.Shapes.AddShape(
                106,
                $calloutLeft,
                $calloutTop,
                $calloutWidth,
                $calloutHeight
            )
            $callout.Name = "SpeechBubble_Callout"
            # Place the native pointer near the lower-right and give it enough
            # reach to meet the top of the mascot.
            $callout.Adjustments.Item(1) = 0.45
            $callout.Adjustments.Item(2) = 0.85

            $callout.TextFrame.TextRange.Text = $bubbleText
            $callout.TextFrame.WordWrap = -1
            $callout.TextFrame.AutoSize = 0
            $callout.TextFrame.VerticalAnchor = 3
            $textRange = $callout.TextFrame.TextRange
            $textRange.ParagraphFormat.Alignment = 2
            $textRange.Font.Name = "Bahnschrift SemiBold"
            $textRange.Font.Size = $fontSize
            $textRange.Font.Bold = -1
            $textRange.Font.Color.RGB = 0xFFFFFF
        }

        # Keep the rebuilt callout in the requested final position on reruns.
        $callout.Width = $calloutWidths[$slideNumber] * $pointsPerInch
        $callout.Height = $calloutHeights[$slideNumber] * $pointsPerInch
        $callout.Left = $pointerTargetX - ($pointerFractionX * $callout.Width)
        $callout.Top = (
            $pointerTargetY -
            ($pointerExtensionY * $callout.Height) -
            $callout.Height
        )
        $callout.Fill.Solid()
        $callout.Fill.ForeColor.RGB = 0x5B5BFF
        $callout.Fill.Transparency = 0
        $callout.Line.ForeColor.RGB = 0xFFFFFF
        $callout.Line.Weight = 1.2
        $callout.Adjustments.Item(1) = 0.45
        $callout.Adjustments.Item(2) = 0.85
        $callout.TextFrame.MarginLeft = 0.14 * $pointsPerInch
        $callout.TextFrame.MarginRight = 0.14 * $pointsPerInch
        $callout.TextFrame.MarginTop = 0.04 * $pointsPerInch
        $callout.TextFrame.MarginBottom = 0.04 * $pointsPerInch

        $mascot = $slide.Shapes.Item("Mascot")
        $mascot.Left = $mascotX
        $mascot.Top = $mascotY
        $mascot.Width = $mascotSize
        $mascot.Height = $mascotSize

        # Native Zoom scales around the group's bounding-box center. Add a
        # completely invisible frame whose bounds are symmetric around the
        # mascot center and contain both visible shapes. This makes the actual
        # group center equal the desired animation pivot.
        $visibleLeft = [Math]::Min($callout.Left, $mascot.Left)
        $visibleTop = [Math]::Min($callout.Top, $mascot.Top)
        $visibleRight = [Math]::Max(
            $callout.Left + $callout.Width,
            $mascot.Left + $mascot.Width
        )
        $visibleBottom = [Math]::Max(
            $callout.Top + $callout.Height,
            $mascot.Top + $mascot.Height
        )
        $pivotHalfWidth = [Math]::Max(
            $mascotOriginX - $visibleLeft,
            $visibleRight - $mascotOriginX
        )
        $pivotHalfHeight = [Math]::Max(
            $mascotOriginY - $visibleTop,
            $visibleBottom - $mascotOriginY
        )
        $pivotFrame = $slide.Shapes.AddShape(
            1,
            $mascotOriginX - $pivotHalfWidth,
            $mascotOriginY - $pivotHalfHeight,
            2.0 * $pivotHalfWidth,
            2.0 * $pivotHalfHeight
        )
        $pivotFrame.Name = "Mascot_Pivot_Frame"
        $pivotFrame.Fill.Visible = 0
        $pivotFrame.Line.Visible = 0

        $overlay = $slide.Shapes.Range($groupParts).Group()
        $overlay.Name = "SpeechOverlay_Group"
        $overlay.ZOrder(0)

        # Use PowerPoint's native Zoom entrance. This is a single click-triggered
        # effect whose scale interpolation is handled by PowerPoint itself.
        # msoAnimEffectZoom=23, msoAnimateLevelNone=0,
        # msoAnimTriggerOnPageClick=1.
        $effect = $sequence.AddEffect($overlay, 23, 0, 1)
        $effect.Timing.Duration = $animationDuration
        $effect.Timing.TriggerType = 1
        $effect.Timing.Accelerate = 0.12
        $effect.Timing.Decelerate = 0.18

    }

    $presentation.Save()
    Write-Output "Applied mascot-centered Zoom entrances to slides 1, 4, 5, 10, and 11."
}
finally {
    if ($null -ne $presentation) {
        $presentation.Close()
    }
    if ($null -ne $powerPoint) {
        $powerPoint.Quit()
    }
}

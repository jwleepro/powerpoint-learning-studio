# PPT Coach - TDD Implementation Plan

## Test-Driven Development Plan

This plan follows Kent Beck's TDD cycle: Red → Green → Refactor. Each test should be implemented one at a time, made to pass, then refactored before moving to the next test.
When running in Claude Code, do not execute in plan mode. Instead, run with "accept edit on" (shift+tab to cycle). It is recommended to run in Sonnet 4.5 or higher model.

---

## Phase 1: PowerPoint COM Interop (FR-101 ~ FR-107)

### 1.1 PPT Instance Connection

- [x] Test: Should detect if PowerPoint is installed on the system
- [x] Test: Should return null when PowerPoint is not running
- [x] Test: Should connect to running PowerPoint instance
- [x] Test: Should handle multiple PowerPoint instances
- [x] Test: Should throw exception when COM connection fails

### 1.2 New Document Creation

- [x] Test: Should create new PowerPoint presentation
- [x] Test: Should return presentation object after creation
- [x] Test: Should create presentation with default blank slide
- [x] Test: Should handle creation failure gracefully

### 1.3 Slide Information Reading

- [x] Test: Should get current slide number
- [x] Test: Should return 0 when no slide is selected
- [x] Test: Should get total slide count
- [x] Test: Should read slide title text
- [x] Test: Should read all text content from slide
- [x] Test: Should get all shapes on slide
- [x] Test: Should return shape type (text, image, table, etc.)

### 1.4 Slide Manipulation

- [x] Test: Should add new blank slide
- [x] Test: Should add slide with specific layout
- [x] Test: Should delete slide by index
- [x] Test: Should move slide from one position to another
- [x] Test: Should prevent deleting when only one slide exists

### 1.5 Element Property Reading

- [x] Test: Should get font name from text shape
- [x] Test: Should get font size from text shape
- [x] Test: Should get font color (RGB) from text shape
- [x] Test: Should get shape fill color
- [x] Test: Should get shape position (Left, Top)
- [x] Test: Should get shape size (Width, Height)
- [x] Test: Should get table cell content
- [x] Test: Should handle shapes without text gracefully

### 1.6 Event Detection

- [x] Test: Should detect slide change event
- [x] Test: Should detect shape selection change event
- [ ] Test: Should detect presentation save event
- [ ] Test: Should detect presentation close event
- [ ] Test: Should unsubscribe from events properly

---

## Phase 2: Step-by-Step Guide System (FR-201 ~ FR-207)

### 2.1 Guide Template Management

- [ ] Test: Should load guide template from JSON file
- [ ] Test: Should parse template with multiple steps
- [ ] Test: Should validate template structure
- [ ] Test: Should return error for invalid template format
- [ ] Test: Should list all available templates
- [ ] Test: Should select template by name

### 2.2 Step Definition

- [ ] Test: Should define step with title and description
- [ ] Test: Should define step with checkpoint conditions
- [ ] Test: Should assign step to specific slide number
- [ ] Test: Should define required actions for step
- [ ] Test: Should define validation rules for step

### 2.3 Progress Status Display

- [ ] Test: Should calculate current step number
- [ ] Test: Should calculate total steps count
- [ ] Test: Should calculate progress percentage
- [ ] Test: Should mark step as completed
- [ ] Test: Should mark step as in-progress
- [ ] Test: Should list all completed steps
- [ ] Test: Should list all pending steps

### 2.4 Step Navigation

- [ ] Test: Should move to next step
- [ ] Test: Should move to previous step
- [ ] Test: Should jump to specific step by index
- [ ] Test: Should prevent moving next when current step incomplete
- [ ] Test: Should allow skipping step with confirmation
- [ ] Test: Should return to first step (restart)

### 2.5 Step Description

- [ ] Test: Should get text description for current step
- [ ] Test: Should get image path for step illustration
- [ ] Test: Should format description with markdown
- [ ] Test: Should provide detailed instructions list

### 2.6 Example Display

- [ ] Test: Should load example image for step
- [ ] Test: Should provide example presentation file path
- [ ] Test: Should show before/after comparison
- [ ] Test: Should handle missing example gracefully

### 2.7 Progress State Persistence

- [ ] Test: Should save progress to file
- [ ] Test: Should load progress from file
- [ ] Test: Should restore guide state on restart
- [ ] Test: Should handle corrupted progress file
- [ ] Test: Should clear progress when starting new guide

---

## Phase 3: Overlay Highlight System (FR-301 ~ FR-307)

### 3.1 Overlay Window

- [ ] Test: Should create transparent overlay window
- [ ] Test: Should position overlay above PowerPoint window
- [ ] Test: Should make overlay always-on-top
- [ ] Test: Should handle multiple monitors
- [ ] Test: Should hide overlay when PowerPoint is minimized

### 3.2 Area Highlighting

- [ ] Test: Should draw rectangle highlight at coordinates
- [ ] Test: Should draw highlight with red border
- [ ] Test: Should draw highlight with semi-transparent fill
- [ ] Test: Should highlight ribbon menu button
- [ ] Test: Should highlight specific slide area
- [ ] Test: Should clear previous highlight before drawing new one

### 3.3 Arrow/Pointer

- [ ] Test: Should draw arrow pointing to target location
- [ ] Test: Should calculate arrow direction automatically
- [ ] Test: Should draw arrow with animation effect
- [ ] Test: Should hide arrow on user click

### 3.4 Tooltip/Callout

- [ ] Test: Should display text tooltip near highlight
- [ ] Test: Should position tooltip to avoid screen edges
- [ ] Test: Should auto-size tooltip based on text length
- [ ] Test: Should style tooltip with readable font

### 3.5 Window Position Sync

- [ ] Test: Should track PowerPoint window position
- [ ] Test: Should update overlay when PPT window moves
- [ ] Test: Should update overlay when PPT window resizes
- [ ] Test: Should handle PPT window minimize/restore
- [ ] Test: Should sync at 60 FPS or better

### 3.6 Click-Through

- [ ] Test: Should allow click-through on non-highlight areas
- [ ] Test: Should pass mouse events to PowerPoint
- [ ] Test: Should detect click on highlighted area
- [ ] Test: Should not block keyboard input to PowerPoint

---

## Phase 4: Real-time Validation & Feedback (FR-401 ~ FR-408)

### 4.1 Font Consistency Check

- [ ] Test: Should detect all font names used in presentation
- [ ] Test: Should count unique font names
- [ ] Test: Should flag inconsistent font usage
- [ ] Test: Should detect all font sizes used
- [ ] Test: Should flag too many different font sizes
- [ ] Test: Should suggest standard font for document

### 4.2 Color Consistency Check

- [ ] Test: Should extract all colors used in shapes
- [ ] Test: Should count unique colors
- [ ] Test: Should flag when more than 3 colors used
- [ ] Test: Should ignore white/black as color count
- [ ] Test: Should suggest color palette

### 4.3 Layout Check

- [ ] Test: Should check title position consistency
- [ ] Test: Should check margin consistency
- [ ] Test: Should check text alignment consistency
- [ ] Test: Should flag misaligned elements
- [ ] Test: Should suggest layout correction

### 4.4 Slide Structure Check

- [ ] Test: Should verify title slide exists
- [ ] Test: Should verify agenda/table of contents exists
- [ ] Test: Should verify required sections exist
- [ ] Test: Should check slide order
- [ ] Test: Should flag missing required slides

### 4.5 Content Check

- [ ] Test: Should detect empty slides
- [ ] Test: Should detect placeholder text not replaced
- [ ] Test: Should detect slides with no title
- [ ] Test: Should detect overly long text (>6 bullets)
- [ ] Test: Should flag missing content

### 4.6 Real-time Notification

- [ ] Test: Should notify immediately when issue detected
- [ ] Test: Should show notification in overlay
- [ ] Test: Should show notification in guide panel
- [ ] Test: Should queue multiple notifications
- [ ] Test: Should allow dismissing notifications

### 4.7 Correction Suggestions

- [ ] Test: Should provide specific correction text
- [ ] Test: Should show before/after preview
- [ ] Test: Should prioritize suggestions by severity
- [ ] Test: Should link suggestion to relevant step

---

## Phase 5: Integration & End-to-End Tests

### 5.1 Complete User Workflow

- [ ] Test: Should complete full guide from start to finish
- [ ] Test: Should handle guide interruption and resume
- [ ] Test: Should save and restore progress correctly
- [ ] Test: Should validate completed presentation passes all checks

### 5.2 Error Handling

- [ ] Test: Should handle PowerPoint crash gracefully
- [ ] Test: Should handle user closing PowerPoint during guide
- [ ] Test: Should handle network/file access errors
- [ ] Test: Should provide meaningful error messages

### 5.3 Performance

- [ ] Test: Should render overlay updates within 16ms
- [ ] Test: Should not block PowerPoint operations
- [ ] Test: Should use less than 100MB memory
- [ ] Test: Should handle large presentations (50+ slides)

### 5.4 Compatibility

- [ ] Test: Should work with PowerPoint 2016
- [ ] Test: Should work with PowerPoint 2019
- [ ] Test: Should work with PowerPoint 2021
- [ ] Test: Should detect unsupported PowerPoint version

---

## Notes

- Mark tests with [x] when complete
- Write one test at a time
- Make it pass with minimal code
- Refactor only when in Green state
- Commit structural changes separately from behavioral changes
- Run all tests before committing

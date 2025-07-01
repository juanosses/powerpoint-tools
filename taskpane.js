// Global arrays to store copied styles, applied names, and colors
let copiedStyles = [];
let appliedNames = [];
let selectedStyleIndex = -1;
let colorList = [];
let positionList = []; // Add this to initialize the position list
let positionCounter = 1; // Keep track of how many positions have been copied
let selectedPositionIndex = -1; // Index for selecting position in the position list
let readyForDesignCount = 0;
let designedCount = 0;
let unmarkedCount = 0;
let selectedNameIndex = -1;
let selectedColorIndex = -1;

// Office.js setup
Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    $(document).ready(function() {
      // Bind functions to buttons
      setupToggleWithArrow("toggleTextStyles", "textStylesSection");
      setupToggleWithArrow("togglePosition", "positionSection");
      setupToggleWithArrow("toggleColors", "colorsSection");
      setupToggleWithArrow("toggleColorChange", "colorChangeSection");
      $("#applyObjectName").on("click", applyObjectName);
      $("#addNameToList").on("click", addNameToList);
      $("#copyAttributes").on("click", copyAttributes);
      $("#applyAttributes").on("click", applyAttributes);
      // Aplica estilo automáticamente cuando se selecciona un nombre en el dropdown
      $("#nameDropdown_textStyle").on("change", function() {
        const selectedName = $(this).val();
        if (selectedName) {
          applyStyleToObjects(selectedName);
        }
      });

      $("#nameDropdown_position").on("change", function() {
        const selectedName = $(this).val();
        if (selectedName) applyPositionToNamedObjects(selectedName);
      });

      $("#nameDropdown_colorText").on("change", function() {
        const selectedName = $(this).val();
        if (selectedName) applyColorToNamedTextBoxes(selectedName);
      });

      $("#nameDropdown_colorFill").on("change", function() {
        const selectedName = $(this).val();
        if (selectedName) applyFillColorToNamedObjects(selectedName);
      });

      $("#nameDropdown_colorStroke").on("change", function() {
        const selectedName = $(this).val();
        if (selectedName) applyLineColorToNamedObjects(selectedName);
      });

      $("#applyStyleToObject").on("click", function() {
        const selectedName = $("#nameDropdown_textStyle").val();
        if (!selectedName) {
          alert("Please select a name from the dropdown.");
          return;
        }
        applyStyleToObjects(selectedName); // ← pasa el nombre como argumento
      });

      $("#addColor").on("click", addColorToList);
      $("#performColorChangeButton").on("click", handleColorChange); // Deck-wide color change
      $("#performSlideColorChangeButton").on("click", handleSlideColorChange); // Active slide color change

      //Delete styles functions
      $("#deleteSelectedStyle").on("click", deleteSelectedStyle);
      $("#deleteSelectedName").on("click", deleteSelectedName);
      $("#deleteSelectedColor").on("click", deleteSelectedColor);
      $("#deleteSelectedPosition").on("click", deleteSelectedPosition);

      // Handle Save button click
      $("#saveStylesButton").on("click", function() {
        saveStylesToLocal();
        console.log("Styles saved by user");
      });

      // Handle Load button click
      $("#loadStylesButton").on("click", function() {
        loadStylesFromLocal();
        console.log("Styles loaded by user");
      });

      // Color text
      $("#applyColorToTextInTextBox").on("click", applyColorToTextInTextBox);
      $("#applyColorToNamedTextBoxes").on("click", function() {
        const selectedName = $("#nameDropdown_colorText").val();
        if (!selectedName) {
          alert("Please select a name from the dropdown.");
          return;
        }
        applyColorToNamedTextBoxes(selectedName);
      });

      // Add plain JavaScript event listener for the manual update button
      document.getElementById("manualUpdateUnmarkedCount").addEventListener("click", function() {
        countUnmarkedSlides();
      });

      // Position-related buttons
      $("#copyPosition").on("click", copyPosition);
      $("#applyPosition").on("click", applyPosition);
      $("#applyPositionToObjects").on("click", function() {
        const selectedName = $("#nameDropdown_position").val();
        if (!selectedName) {
          alert("Please select a name from the dropdown.");
          return;
        }
        applyPositionToNamedObjects(selectedName);
      });

      $("#applyLineColorToObject").on("click", function() {
        const selectedName = $("#nameDropdown_colorStroke").val();
        if (!selectedName) {
          alert("Please select a name from the dropdown.");
          return;
        }
        applyLineColorToNamedObjects(selectedName);
      });

      $("#applyFillColorToObject").on("click", function() {
        const selectedName = $("#nameDropdown_colorFill").val();
        if (!selectedName) {
          alert("Please select a name from the dropdown.");
          return;
        }
        applyFillColorToNamedObjects(selectedName);
      });

      // Initialize counters
      updateCounts();

      // Bind functions to tag buttons
      $("#markReadyForDesign").on("click", () => {
        markSlide("Ready for Design", "orange");
      });

      $("#markDesigned").on("click", () => {
        markSlide("Designed", "green");
      });

      // Count unmarked slides at load
      countUnmarkedSlides();

      // Listeners for selecting items in the lists
      $("#nameList").on("click", ".name-item", function() {
        $(".name-item").removeClass("selected");
        $(this).addClass("selected");
      });

      $("#styleList").on("click", ".style-item", function() {
        $(".style-item").removeClass("selected");
        $(this).addClass("selected");
        selectedStyleIndex = $(this).data("index");
      });

      $("#positionList").on("click", ".position-item", function() {
        $(".position-item").removeClass("selected");
        $(this).addClass("selected");
        selectedPositionIndex = $(this).data("index");
      });

      $("#colorList").on("click", ".color-item", function() {
        $(".color-item").removeClass("selected");
        $(this).addClass("selected");
      });

      $("#applyFillColor").on("click", applyFillColor);
      $("#applyLineColor").on("click", applyLineColor);

      // Toggle Naming section
      $("#toggleNaming").on("click", function() {
        $("#namingSection").slideToggle(); // Use jQuery's slideToggle for animation
      });

      // Toggle Text Styles section
      $("#toggleTextStyles").on("click", function() {
        $("#textStylesSection").slideToggle();
      });

      // Toggle Colors section
      $("#toggleColors").on("click", function() {
        $("#colorsSection").slideToggle();
      });

      // Toggle Position section
      $("#togglePosition").on("click", function() {
        $("#positionSection").slideToggle();
      });

      // Toggle the color change section
      $("#toggleColorChange").on("click", function() {
        $("#colorChangeSection").slideToggle(); // Use jQuery's slideToggle for animation
      });
    });
  }
});

// Add Name to List
function addNameToList() {
  const name = $("#IDnameInput")
    .val()
    .trim();
  if (name && !appliedNames.includes(name)) {
    appliedNames.push(name);
    updateNameList();
    $("#IDnameInput").val(""); // Clear the input after adding
    console.log("Name added to list successfully:", name);
  } else {
    console.log("Please enter a unique name or select a different name.");
  }
}

// Apply Name to Selected Shapes
function applyObjectName() {
  const selectedName = $(".name-item.selected")
    .text()
    .trim();
  if (selectedName) {
    PowerPoint.run(async (context) => {
      const selectedObjects = context.presentation.getSelectedShapes();
      selectedObjects.load("items/name");
      await context.sync();

      selectedObjects.items.forEach(function(shape) {
        shape.name = selectedName; // Apply the selected name to each shape
      });
      await context.sync();
      console.log("Name applied successfully to selected objects!");
    }).catch((error) => {
      console.error("Error applying name: " + error);
    });
  } else {
    console.log("No name selected. Please select a name from the list.");
  }
}

function updateNameDropdowns() {
  const dropdowns = [
    "#nameDropdown_textStyle",
    "#nameDropdown_position",
    "#nameDropdown_colors",
    "#nameDropdown_colorSwap",
    "#nameDropdown_colorText",
    "#nameDropdown_colorFill",
    "#nameDropdown_colorStroke"
  ];

  dropdowns.forEach((selector) => {
    const dropdown = $(selector);
    dropdown.empty();
    dropdown.append(`<option value="" hidden selected>Apply to ID</option>`);

    appliedNames.forEach((name) => {
      dropdown.append(`<option value="${name}">${name}</option>`);
    });
  });
}

// Function to update the name list and handle selection
function updateNameList() {
  const nameList = $("#nameList");
  nameList.empty();
  appliedNames.forEach((name, index) => {
    nameList.append(`
      <li class="name-item" data-index="${index}">${name}</li>
    `);
  });

  // Handle name item selection
  $(".name-item").on("click", function() {
    $(".name-item").removeClass("selected");
    $(this).addClass("selected");
    selectedNameIndex = $(this).data("index");
  });

  // 🔄 Actualizar el dropdown después de actualizar la lista
  updateNameDropdowns();
}

// Function to delete the selected name
function deleteSelectedName() {
  if (selectedNameIndex >= 0 && selectedNameIndex < appliedNames.length) {
    appliedNames.splice(selectedNameIndex, 1); // Remove the selected name
    selectedNameIndex = -1; // Reset selection
    updateNameList(); // Refresh the UI
    saveStylesToLocal(); // Save the updated data
    console.log("Selected name deleted");
  } else {
    alert("No name selected");
  }
}

// Copy Text Attributes from Selected Text
function copyAttributes() {
  const styleName = $("#styleNameInput")
    .val()
    .trim(); // Get the style name from input

  if (!styleName) {
    alert("Please enter a name for the style.");
    return; // Stop if no style name is provided
  }

  PowerPoint.run(async (context) => {
    const selectedShape = context.presentation.getSelectedShapes();
    selectedShape.load("items/name,items/textFrame");
    await context.sync();

    // Ensure at least one shape is selected
    if (selectedShape.items.length === 0) {
      alert("No text box selected.");
      return;
    }

    // Assuming we work with the first selected shape
    let shape = selectedShape.items[0];

    // Load all necessary text and paragraph properties
    shape.textFrame.load("textRange/font, textRange/paragraphFormat");
    shape.textFrame.load("topMargin, bottomMargin, leftMargin, rightMargin");
    await context.sync(); // Ensure properties are loaded

    // Load paragraph format properties separately (including LineRuleWithin)
    shape.textFrame.textRange.paragraphFormat.load("horizontalAlignment, lineSpacing, lineRuleWithin");
    await context.sync(); // Sync again to ensure these properties are available

    let lineSpacingValue = shape.textFrame.textRange.paragraphFormat.lineSpacing;
    let lineRuleWithin = shape.textFrame.textRange.paragraphFormat.lineRuleWithin;

    console.log(`Retrieved line spacing for shape "${shape.name}": ${lineSpacingValue}`);
    console.log(`Retrieved LineRuleWithin for shape "${shape.name}": ${lineRuleWithin}`);

    // If `lineSpacing` is undefined but `lineRuleWithin` is set to multiple, we extract its value
    if ((lineSpacingValue === undefined || lineSpacingValue === null) && lineRuleWithin) {
      console.warn(`Using LineRuleWithin value as line spacing: ${lineRuleWithin}`);
      lineSpacingValue = lineRuleWithin;
    }

    const newStyle = {
      styleName: styleName, // Include the user-defined style name
      fontName: shape.textFrame.textRange.font.name,
      fontSize: shape.textFrame.textRange.font.size,
      fontColor: shape.textFrame.textRange.font.color,
      fontWeight: shape.textFrame.textRange.font.bold ? "bold" : "normal",
      fontItalic: shape.textFrame.textRange.font.italic ? "italic" : "normal",
      underline: shape.textFrame.textRange.font.underline,

      // Paragraph alignment
      alignment: shape.textFrame.textRange.paragraphFormat.horizontalAlignment,

      // Line Spacing (FIXED)
      lineSpacing: lineSpacingValue,

      // Margins
      topMargin: shape.textFrame.topMargin,
      bottomMargin: shape.textFrame.bottomMargin,
      leftMargin: shape.textFrame.leftMargin,
      rightMargin: shape.textFrame.rightMargin
    };

    console.log(`Copied Style from "${shape.name}":`, newStyle);

    copiedStyles.push(newStyle); // Store the new style with its name
    updateStyleList(); // Update the style list with the new style
    $("#styleNameInput").val(""); // Clear the style name input after use
    console.log("Attributes, including line spacing and alignment, copied successfully!");
  }).catch((error) => {
    console.error("Error copying attributes, including line spacing: " + error);
  });
}

// Function to update the style list and handle selection
function updateStyleList() {
  const styleList = $("#styleList");
  styleList.empty();
  copiedStyles.forEach((style, index) => {
    styleList.append(`
      <li class="style-item" data-index="${index}">
        ${style.styleName}, ${style.fontName}, ${style.fontSize}px
      </li>
    `);
  });

  // Handle style item selection
  $(".style-item").on("click", function() {
    $(".style-item").removeClass("selected");
    $(this).addClass("selected");
    selectedStyleIndex = $(this).data("index");
  });
}

// Function to delete the selected style
function deleteSelectedStyle() {
  if (selectedStyleIndex >= 0 && selectedStyleIndex < copiedStyles.length) {
    copiedStyles.splice(selectedStyleIndex, 1); // Remove the selected style
    selectedStyleIndex = -1; // Reset selection
    updateStyleList(); // Refresh the UI
    saveStylesToLocal(); // Save the updated data
    console.log("Selected style deleted");
  } else {
    alert("No style selected");
  }
}

// Add Fill Color to the Color List
function addColorToList() {
  const colorName = $("#styleNameInput")
    .val()
    .trim();

  if (!colorName) {
    alert("Please enter a name for the color.");
    return;
  }

  PowerPoint.run(async (context) => {
    let selectedShape = context.presentation.getSelectedShapes();
    selectedShape.load("fill");
    await context.sync();

    let fillColor = selectedShape.items[0].fill.foregroundColor;

    if (!colorList.some((item) => item.color === fillColor && item.name === colorName)) {
      colorList.push({ name: colorName, color: fillColor });
      updateColorList();
      $("#styleNameInput").val(""); // Clear input after adding
      console.log(`Color "${colorName}" added to list with value: ${fillColor}`);
    } else {
      alert("That color with this name already exists in the list.");
    }
  }).catch((error) => {
    console.error("Error adding color: " + error);
  });
}

// Function to update the color list and handle selection
function updateColorList() {
  const colorListContainer = $("#colorList");
  colorListContainer.empty();
  colorList.forEach((color, index) => {
    colorListContainer.append(`
      <li class="color-item" data-index="${index}" style="background-color:${color.color};">${color.color}</li>
    `);
  });

  // Handle color item selection
  $(".color-item").on("click", function() {
    $(".color-item").removeClass("selected");
    $(this).addClass("selected");
    selectedColorIndex = $(this).data("index");
  });
}

// Function to delete the selected color
function deleteSelectedColor() {
  if (selectedColorIndex >= 0 && selectedColorIndex < colorList.length) {
    colorList.splice(selectedColorIndex, 1); // Remove the selected color
    selectedColorIndex = -1; // Reset selection
    updateColorList(); // Refresh the UI
    saveStylesToLocal(); // Save the updated data
    console.log("Selected color deleted");
  } else {
    alert("No color selected");
  }
}

// Apply Attributes to Selected Text
function applyAttributes() {
  if (selectedStyleIndex >= 0 && selectedStyleIndex < copiedStyles.length) {
    let selectedStyle = copiedStyles[selectedStyleIndex];
    PowerPoint.run(async (context) => {
      const selectedShape = context.presentation.getSelectedShapes();
      selectedShape.load(
        "items/textFrame/topMargin, items/textFrame/bottomMargin, items/textFrame/leftMargin, items/textFrame/rightMargin, items/textFrame/textRange/font, items/textFrame/textRange/paragraphFormat"
      );
      await context.sync();

      // Ensure at least one shape is selected
      if (selectedShape.items.length === 0) {
        alert("No text box selected.");
        return;
      }

      // Assuming we work with the first selected shape
      let shape = selectedShape.items[0];

      // Log the copied line spacing before applying it
      console.log(`Applying line spacing: ${selectedStyle.lineSpacing}`);

      // Apply the font attributes
      shape.textFrame.textRange.font.name = selectedStyle.fontName;
      shape.textFrame.textRange.font.size = selectedStyle.fontSize;
      shape.textFrame.textRange.font.color = selectedStyle.fontColor;
      shape.textFrame.textRange.font.bold = selectedStyle.fontWeight === "bold";
      shape.textFrame.textRange.font.italic = selectedStyle.fontItalic === "italic";
      shape.textFrame.textRange.font.underline = selectedStyle.underline;

      // Apply the paragraph alignment
      shape.textFrame.textRange.paragraphFormat.horizontalAlignment = selectedStyle.alignment || "Left"; // Default to "Left" if undefined

      // Apply line spacing (FIXED)
      if (selectedStyle.lineSpacing !== "Not Set") {
        shape.textFrame.textRange.paragraphFormat.lineSpacing = selectedStyle.lineSpacing;
        console.log(`Line spacing applied: ${selectedStyle.lineSpacing}`);
      } else {
        console.log("Line spacing was not set in the copied style.");
      }

      // Apply the margins
      shape.textFrame.topMargin = selectedStyle.topMargin || 0;
      shape.textFrame.bottomMargin = selectedStyle.bottomMargin || 0;
      shape.textFrame.leftMargin = selectedStyle.leftMargin || 0;
      shape.textFrame.rightMargin = selectedStyle.rightMargin || 0;

      await context.sync();
      console.log("Attributes and line spacing applied successfully!");
    }).catch((error) => {
      console.error("Error applying attributes, including line spacing: " + error);
    });
  } else {
    console.log("No style selected or index out of range.");
  }
}

// Function to apply fill color
function applyFillColor() {
  const selectedColorItem = $(".color-item.selected");
  if (!selectedColorItem.length) {
    alert("No color selected to apply.");
    return;
  }

  // Extract the background color in a format that Office JS expects (e.g., #RRGGBB)
  const selectedColor = rgbToHex(selectedColorItem.css("background-color"));
  if (!selectedColor) {
    alert("Invalid color format.");
    return;
  }

  PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items/fill");

    await context.sync();

    selectedShapes.items.forEach((shape) => {
      shape.fill.setSolidColor(selectedColor);
    });

    await context.sync();
    alert("Fill color applied successfully!");
  }).catch(function(error) {
    console.error("Error applying fill color:", error);
    alert("Error applying fill color.");
  });
}

function rgbToHex(rgb) {
  if (!rgb || !rgb.startsWith("rgb")) return rgb; // Return the original if it's not RGB format
  const parts = rgb.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*\d+)?\)$/);
  if (!parts) return null;
  const r = parseInt(parts[1], 10)
    .toString(16)
    .padStart(2, "0");
  const g = parseInt(parts[2], 10)
    .toString(16)
    .padStart(2, "0");
  const b = parseInt(parts[3], 10)
    .toString(16)
    .padStart(2, "0");
  return `#${r}${g}${b}`;
}

// Function to apply line color
function applyLineColor() {
  const selectedColorItem = $(".color-item.selected");
  if (!selectedColorItem.length) {
    alert("No color selected to apply.");
    return;
  }

  // Extract the background color in a format that Office JS expects (e.g., #RRGGBB)
  const selectedColor = rgbToHex(selectedColorItem.css("background-color"));
  if (!selectedColor) {
    alert("Invalid color format.");
    return;
  }

  PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items/lineFormat"); // Load the lineFormat property of the selected shapes
    await context.sync();

    selectedShapes.items.forEach((shape) => {
      if (shape.lineFormat) {
        shape.lineFormat.color = selectedColor; // Set the border color of the shape
        console.log(`Line color ${selectedColor} applied successfully to shape.`);
      } else {
        console.log("Selected shape does not support line formatting or line formatting properties are not available.");
      }
    });

    await context.sync();
    alert("Line color applied successfully!");
  }).catch((error) => {
    console.error("Error applying line color:", error);
    alert("Error applying line color.");
  });
}

function applyStyleToObjects(selectedName) {
  if (!selectedName) {
    console.log("No name received. Please select a name.");
    return;
  }

  if (selectedStyleIndex >= 0 && selectedStyleIndex < copiedStyles.length) {
    let selectedStyle = copiedStyles[selectedStyleIndex];

    PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides.load("items");
        await context.sync();

        console.log(`Applying style to objects named: "${selectedName}"`);
        console.log("Using style:", selectedStyle);

        for (let slide of slides.items) {
          const shapes = slide.shapes.load("items/name");
          await context.sync();

          for (let shape of shapes.items) {
            console.log(`Shape: "${shape.name}" vs Selected: "${selectedName}"`);

            if (shape.name === selectedName) {
              console.log(`🔍 Match found: "${shape.name}", attempting to apply style...`);

              try {
                shape.textFrame.load("textRange/font, textRange/paragraphFormat, topMargin, bottomMargin, leftMargin, rightMargin");
                await context.sync();

                if (shape.textFrame && shape.textFrame.textRange) {
                  const textRange = shape.textFrame.textRange;

                  textRange.font.name = selectedStyle.fontName;
                  textRange.font.size = selectedStyle.fontSize;
                  textRange.font.color = selectedStyle.fontColor;
                  textRange.font.bold = selectedStyle.fontWeight === "bold";
                  textRange.font.italic = selectedStyle.fontItalic === "italic";
                  textRange.font.underline = selectedStyle.underline;

                  textRange.paragraphFormat.horizontalAlignment = selectedStyle.alignment || "Left";
                  shape.textFrame.topMargin = selectedStyle.topMargin || 0;
                  shape.textFrame.bottomMargin = selectedStyle.bottomMargin || 0;
                  shape.textFrame.leftMargin = selectedStyle.leftMargin || 0;
                  shape.textFrame.rightMargin = selectedStyle.rightMargin || 0;

                  await context.sync();
                  console.log(`✅ Style applied to shape: "${shape.name}"`);
                } else {
                  console.warn(`⚠️ Shape "${shape.name}" does not support text formatting.`);
                }
              } catch (shapeError) {
                console.warn(`⚠️ Skipped shape "${shape.name}":`, shapeError.message);
              }
            }
          }
        }

        console.log("🎯 Style application complete.");
      } catch (error) {
        console.error("❌ Error applying style to objects:", error);
      }
    }).catch((error) => {
      console.error("❌ PowerPoint.run failed:", error);
    });
  } else {
    console.log("No style selected.");
  }
}

function applyLineColorToNamedObjects(selectedName) {
  const selectedColorItem = $(".color-item.selected");
  selectedName =
    selectedName ||
    $(".name-item.selected")
      .text()
      .trim();

  let lineColor = selectedColorItem.css("background-color");

  if (!selectedName || !selectedColorItem.length) {
    console.log("No name or color selected.");
    return;
  }

  lineColor = rgbToHex(lineColor);
  if (!/^#[0-9A-F]{6}$/i.test(lineColor)) {
    console.log("Invalid color format.");
    return;
  }

  PowerPoint.run(async (context) => {
    const slides = context.presentation.slides.load("items");
    await context.sync();

    for (let slide of slides.items) {
      const shapes = slide.shapes.load("items/name,items/lineFormat");
      await context.sync();

      for (let shape of shapes.items) {
        if (shape.name === selectedName && shape.lineFormat) {
          shape.lineFormat.color = lineColor;
          await context.sync();
          console.log(`Line color applied successfully to object named "${selectedName}"!`);
        }
      }
    }
  }).catch((error) => {
    console.error("Error applying line color to objects: " + error);
  });
}

function applyFillColorToNamedObjects(selectedName) {
  const selectedColorItem = $(".color-item.selected");
  selectedName =
    selectedName ||
    $(".name-item.selected")
      .text()
      .trim();

  let fillColor = selectedColorItem.css("background-color");

  if (!selectedName || !selectedColorItem.length) {
    console.log("No name or color selected.");
    return;
  }

  fillColor = rgbToHex(fillColor);
  if (!/^#[0-9A-F]{6}$/i.test(fillColor)) {
    console.log("Invalid color format.");
    return;
  }

  PowerPoint.run(async (context) => {
    const slides = context.presentation.slides.load("items");
    await context.sync();

    for (let slide of slides.items) {
      const shapes = slide.shapes.load("items/name,items/fill");
      await context.sync();

      for (let shape of shapes.items) {
        if (shape.name === selectedName && shape.fill) {
          shape.fill.setSolidColor(fillColor);
          await context.sync();
          console.log(`Fill color applied successfully to object named "${selectedName}"!`);
        }
      }
    }
  }).catch((error) => {
    console.error("Error applying fill color to objects: " + error);
  });
}

function copyPosition() {
  const positionName = $("#styleNameInput")
    .val()
    .trim(); // Get the position name entered by the user

  if (!positionName) {
    alert("Please enter a name for the position.");
    return; // Don't proceed if no name is entered
  }

  PowerPoint.run(async (context) => {
    const selectedShape = context.presentation.getSelectedShapes();
    selectedShape.load("items/left,items/top");
    await context.sync();

    let shape = selectedShape.items[0];

    const newPosition = {
      name: positionName, // Store only the name entered by the user
      left: shape.left,
      top: shape.top
    };

    positionList.push(newPosition);
    updatePositionList();
    $("#styleNameInput").val(""); // Clear the input after adding
  }).catch((error) => {
    console.error("Error copying position: " + error);
  });
}

// Function to update the position list and handle selection
function updatePositionList() {
  const positionListContainer = $("#positionList");
  positionListContainer.empty();
  positionList.forEach((position, index) => {
    positionListContainer.append(`
      <li class="position-item" data-index="${index}">${position.name}</li>
    `);
  });

  // Handle position item selection
  $(".position-item").on("click", function() {
    $(".position-item").removeClass("selected");
    $(this).addClass("selected");
    selectedPositionIndex = $(this).data("index");
  });
}

// Function to delete the selected position
function deleteSelectedPosition() {
  if (selectedPositionIndex >= 0 && selectedPositionIndex < positionList.length) {
    positionList.splice(selectedPositionIndex, 1); // Remove the selected position
    selectedPositionIndex = -1; // Reset selection
    updatePositionList(); // Refresh the UI
    saveStylesToLocal(); // Save the updated data
    console.log("Selected position deleted");
  } else {
    alert("No position selected");
  }
}

// Apply Position to Selected Shape
function applyPosition() {
  if (selectedPositionIndex >= 0 && selectedPositionIndex < positionList.length) {
    const selectedPosition = positionList[selectedPositionIndex];
    PowerPoint.run(async (context) => {
      const selectedShape = context.presentation.getSelectedShapes();
      selectedShape.load("items/left,items/top");
      await context.sync();

      let shape = selectedShape.items[0];
      shape.left = selectedPosition.left;
      shape.top = selectedPosition.top;

      await context.sync();
    }).catch((error) => {
      console.error("Error applying position: " + error);
    });
  }
}

// Apply Position to Objects with Specific Name
function applyPositionToNamedObjects(selectedName) {
  if (selectedPositionIndex >= 0 && selectedPositionIndex < positionList.length && selectedName) {
    const selectedPosition = positionList[selectedPositionIndex];
    PowerPoint.run(async (context) => {
      const slides = context.presentation.slides.load("items");
      await context.sync();

      for (let slide of slides.items) {
        const shapes = slide.shapes.load("items/name,items/left,items/top");
        await context.sync();

        for (let shape of shapes.items) {
          if (shape.name === selectedName) {
            shape.left = selectedPosition.left;
            shape.top = selectedPosition.top;
          }
        }
        await context.sync();
      }
    }).catch((error) => {
      console.error("Error applying position to objects: " + error);
    });
  } else {
    console.log("No position or name selected.");
  }
}

// Function to mark the slide with a label
// Function to mark the slide with a label
function markSlide(label, color) {
  PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items"); // Load the selected slides
    await context.sync(); // Ensure slides are loaded

    // Check if any slide is selected
    if (!selectedSlides || !selectedSlides.items || selectedSlides.items.length === 0) {
      console.error("No slide selected.");
      alert("No slide selected. Please select a slide.");
      return; // Exit the function if no slides are selected
    }

    const slide = selectedSlides.items[0]; // Get the first selected slide
    const shapes = slide.shapes;

    // Remove any existing labels first
    await removeExistingLabels(slide);

    // Add the new label as a text box on the slide
    const textBox = shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle // Add a rectangular shape
    );

    textBox.left = 400; // Left position (7.5 cm in points)
    textBox.top = 0; // Top position
    textBox.width = 212.625; // Width (7.5 cm in points)
    textBox.height = 42.525; // Height (1.5 cm in points)

    // Set the displayed text based on the label
    textBox.textFrame.textRange.text = label === "Designed" ? "Design Locked" : label; // Display "Design Locked" for "Designed"
    textBox.fill.setSolidColor(color); // Set the fill color

    // Set font size and alignment
    textBox.textFrame.textRange.font.size = 16; // Set font size to 16pt
    textBox.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle; // Center vertically

    // Assign internal names for easier identification
    if (label === "Ready for Design") {
      textBox.name = "ReadyForDesign";
    } else if (label === "Designed") {
      textBox.name = "Designed"; // Internal name stays "Designed" for counting
    }

    await context.sync(); // Sync the changes

    // Update counts after marking the slide
    updateCounts();
    countUnmarkedSlides(); // Call to update unmarked slides count
  }).catch((error) => {
    console.error("Error marking slide: " + error);
  });
}

function updateCounts() {
  PowerPoint.run(async (context) => {
    try {
      const slides = context.presentation.slides.load("items"); // Load all slides
      await context.sync(); // Ensure slides are loaded

      // If no slides exist, exit the function
      if (!slides || slides.items.length === 0) {
        console.error("No slides available in the presentation.");
        return;
      }

      // Reset counts before recalculating
      readyForDesignCount = 0;
      designedCount = 0;

      // Loop through slides to check for shapes with specific names
      for (let slide of slides.items) {
        const shapes = slide.shapes.load("items/name"); // Load shape names
        await context.sync(); // Ensure shapes are loaded

        let hasReadyForDesign = false;
        let hasDesigned = false;

        // Check each shape for the names "ReadyForDesign" or "Designed"
        for (let shape of shapes.items) {
          if (shape.name === "ReadyForDesign") {
            hasReadyForDesign = true;
          } else if (shape.name === "Designed") {
            hasDesigned = true;
          }
        }

        if (hasReadyForDesign) {
          readyForDesignCount++;
        }
        if (hasDesigned) {
          designedCount++;
        }
      }

      // Update UI with the counts
      $("#readyForDesignCount").text(readyForDesignCount);
      $("#designedCount").text(designedCount);

      // Call countUnmarkedSlides to update unmarked slides
      countUnmarkedSlides();
    } catch (error) {
      console.error("Error updating counts: " + error);
    }
  });
}

function countUnmarkedSlides() {
  PowerPoint.run(async (context) => {
    try {
      const slides = context.presentation.slides.load("items"); // Load all slides
      await context.sync();

      if (!slides || slides.items.length === 0) {
        console.error("No slides available to count.");
        return;
      }

      let unmarkedSlidesCount = 0;

      // Loop through all slides to check if they are unmarked
      for (let slide of slides.items) {
        const shapes = slide.shapes.load("items/name"); // Load the names of shapes on the slide
        await context.sync();

        let isMarked = false;

        // Check if the slide has any shape with the name "ReadyForDesign" or "Designed"
        for (let shape of shapes.items) {
          if (shape.name === "ReadyForDesign" || shape.name === "Designed") {
            isMarked = true; // If found, mark the slide as "marked"
            break;
          }
        }

        if (!isMarked) {
          unmarkedSlidesCount++; // Count slides that are unmarked
        }
      }

      // Clear the counter display before updating
      $("#unmarkedCount").text("0"); // Reset the counter display

      // Update the counter display with the recalculated unmarked slides count
      $("#unmarkedCount").text(unmarkedSlidesCount);

      console.log("Manually updated unmarked slides count:", unmarkedSlidesCount); // Log for debugging
    } catch (error) {
      console.error("Error counting unmarked slides: " + error);
    }
  });
}

// Function to remove existing labels from the slide
async function removeExistingLabels(slide) {
  slide.shapes.load("items/name"); // Load the names of each shape on the slide
  await slide.context.sync(); // Ensure shapes are loaded

  // Loop through the shapes and remove those with the names "ReadyForDesign" or "Designed"
  for (let shape of slide.shapes.items) {
    if (shape.name === "ReadyForDesign" || shape.name === "Designed") {
      shape.delete(); // Remove the label shape
    }
  }

  await slide.context.sync(); // Sync the changes
}

function saveStylesToLocal() {
  const stylesJson = JSON.stringify(copiedStyles);
  const namesJson = JSON.stringify(appliedNames);
  const positionsJson = JSON.stringify(positionList);
  const colorsJson = JSON.stringify(colorList);

  Office.context.document.settings.set("copiedStyles", stylesJson);
  Office.context.document.settings.set("appliedNames", namesJson);
  Office.context.document.settings.set("positionList", positionsJson);
  Office.context.document.settings.set("colorList", colorsJson);

  Office.context.document.settings.saveAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Data saved successfully!");
    } else {
      console.error("Error saving data:", asyncResult.error.message);
    }
  });
}

function loadStylesFromLocal() {
  const stylesJson = Office.context.document.settings.get("copiedStyles");
  const namesJson = Office.context.document.settings.get("appliedNames");
  const positionsJson = Office.context.document.settings.get("positionList");
  const colorsJson = Office.context.document.settings.get("colorList");

  if (stylesJson) copiedStyles = JSON.parse(stylesJson);
  if (namesJson) appliedNames = JSON.parse(namesJson);
  if (positionsJson) positionList = JSON.parse(positionsJson);
  if (colorsJson) colorList = JSON.parse(colorsJson);

  updateStyleList();
  updateNameList();
  updatePositionList();
  updateColorList();
}

function applyColorToNamedTextBoxes(selectedName) {
  const selectedColorItem = $(".color-item.selected");
  selectedName =
    selectedName ||
    $(".name-item.selected")
      .text()
      .trim();

  if (!selectedColorItem.length) {
    alert("No color selected to apply.");
    return;
  }

  if (!selectedName) {
    alert("No name selected. Please select a name.");
    return;
  }

  const selectedColor = rgbToHex(selectedColorItem.css("background-color"));
  if (!selectedColor) {
    alert("Invalid color format.");
    return;
  }

  PowerPoint.run(async (context) => {
    const slides = context.presentation.slides.load("items");
    await context.sync();

    for (let slide of slides.items) {
      const shapes = slide.shapes.load("items/name");
      await context.sync();

      for (let shape of shapes.items) {
        if (shape.name === selectedName) {
          shape.textFrame.load("textRange/font");
          await context.sync();

          if (shape.textFrame && shape.textFrame.textRange) {
            shape.textFrame.textRange.font.color = selectedColor;
          }
        }
      }

      await context.sync();
    }

    console.log(`Color applied to all text within text boxes named "${selectedName}".`);
  }).catch((error) => {
    console.error("Error applying color to named text boxes:", error);
  });
}

// Function to apply the selected color to the text of the selected text box
function applyColorToTextInTextBox() {
  const selectedColorItem = $(".color-item.selected");
  if (!selectedColorItem.length) {
    alert("No color selected to apply.");
    return;
  }

  // Extract the selected color in a format that Office JS expects (e.g., #RRGGBB)
  const selectedColor = rgbToHex(selectedColorItem.css("background-color"));
  if (!selectedColor) {
    alert("Invalid color format.");
    return;
  }

  PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items/textFrame");

    await context.sync();

    if (selectedShapes.items.length === 0) {
      console.error("No text box selected.");
      alert("Please select a text box.");
      return;
    }

    // Apply the selected color to the font color of each selected text box
    selectedShapes.items.forEach((shape) => {
      if (shape.textFrame && shape.textFrame.textRange) {
        shape.textFrame.textRange.font.color = selectedColor; // Apply color to the text font
      }
    });

    await context.sync();
    console.log("Color applied to the text within the selected text box.");
  }).catch((error) => {
    console.error("Error applying color to text within text box:", error);
  });
}

//Handle deck color change
function handleColorChange() {
  const initialColor = normalizeColor(
    $("#initialColorInput")
      .val()
      .trim()
  );
  const newColor = normalizeColor(
    $("#newColorInput")
      .val()
      .trim()
  );

  console.log(`Normalized Initial Color: ${initialColor}`);
  console.log(`Normalized New Color: ${newColor}`);

  if (initialColor && newColor) {
    changeColorsInPresentation(initialColor, newColor);
  } else {
    alert("Please enter valid colors in Hex or RGB format.");
  }
}

//Swap colors in entire deck
function changeColorsInPresentation(initialColor, newColor) {
  PowerPoint.run(async (context) => {
    try {
      const slides = context.presentation.slides.load("items");
      await context.sync();

      console.log(`Changing all occurrences of "${initialColor}" to "${newColor}" across the presentation.`);

      for (let slide of slides.items) {
        console.log(`Processing slide: ${slide.id}`);

        // Check and change background color only if it matches the initial color
        try {
          if (slide.background && slide.background.fill) {
            slide.background.fill.load("foregroundColor"); // Load background color
            await context.sync();
            const backgroundColor = slide.background.fill.foregroundColor;
            console.log(`Slide ${slide.id} background color: ${backgroundColor}`);
            if (rgbOrHexEquals(backgroundColor, initialColor)) {
              slide.background.fill.setSolidColor(newColor);
              console.log(`Background color updated for slide: ${slide.id}`);
            }
          } else {
            console.log(`Slide ${slide.id} does not support solid background fill.`);
          }
        } catch (bgError) {
          console.warn(`Background color not updated for slide: ${slide.id}. Error: ${bgError.message}`);
        }

        // Load shapes and their properties
        const shapes = slide.shapes.load("items/name,items/fill,items/lineFormat"); // Safely load basic properties
        await context.sync();

        for (let shape of shapes.items) {
          console.log(`Checking shape: "${shape.name}"`);

          // Load and update fill color
          if (shape.fill) {
            shape.fill.load("foregroundColor");
            await context.sync();
            console.log(`Shape "${shape.name}" has fill color: ${shape.fill.foregroundColor}`);
            if (rgbOrHexEquals(shape.fill.foregroundColor, initialColor)) {
              shape.fill.setSolidColor(newColor);
              console.log(`Updated fill color for shape: ${shape.name}`);
            }
          }

          // Load and update line color
          if (shape.lineFormat) {
            shape.lineFormat.load("color");
            await context.sync();
            console.log(`Shape "${shape.name}" has line color: ${shape.lineFormat.color}`);
            if (rgbOrHexEquals(shape.lineFormat.color, initialColor)) {
              shape.lineFormat.color = newColor;
              console.log(`Updated line color for shape: ${shape.name}`);
            }
          }

          // Load and update text color safely
          if (shape.textFrame && shape.textFrame.textRange) {
            try {
              shape.textFrame.textRange.font.load("color");
              await context.sync();
              console.log(`Shape "${shape.name}" has text color: ${shape.textFrame.textRange.font.color}`);
              if (rgbOrHexEquals(shape.textFrame.textRange.font.color, initialColor)) {
                shape.textFrame.textRange.font.color = newColor;
                console.log(`Updated text color for shape: ${shape.name}`);
              }
            } catch (textError) {
              console.warn(`Text color not updated for shape: ${shape.name}. Error: ${textError.message}`);
            }
          }
        }

        await context.sync(); // Sync after each slide
      }

      console.log("Color changes completed.");
    } catch (error) {
      console.error("Error changing colors across the presentation:", error);
    }
  });
}

// Utility function to validate and compare colors (RGB or Hexadecimal)
function rgbOrHexEquals(color1, color2) {
  // Normalize colors to Hexadecimal for comparison
  return normalizeColor(color1) === normalizeColor(color2);
}

// Normalize a color to Hexadecimal
function normalizeColor(color) {
  if (!color) return null;

  if (color.startsWith("#")) {
    return color.toUpperCase(); // Already Hex
  }

  // Handle Hex without a `#` prefix
  if (/^[0-9A-F]{6}$/i.test(color)) {
    return `#${color.toUpperCase()}`;
  }

  // Convert RGB to Hex
  const match = color.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
  if (match) {
    const r = parseInt(match[1])
      .toString(16)
      .padStart(2, "0");
    const g = parseInt(match[2])
      .toString(16)
      .padStart(2, "0");
    const b = parseInt(match[3])
      .toString(16)
      .padStart(2, "0");
    return `#${r}${g}${b}`.toUpperCase();
  }

  return null; // Invalid format
}

// Function to handle color swap
function handleSlideColorChange() {
  const initialColor = normalizeColor(
    $("#initialColorInput")
      .val()
      .trim()
  );
  const newColor = normalizeColor(
    $("#newColorInput")
      .val()
      .trim()
  );

  if (initialColor && newColor) {
    changeColorsOnActiveSlide(initialColor, newColor);
  } else {
    alert("Please enter valid colors in Hex or RGB format.");
  }
}

// Swap color in active slide
function changeColorsOnActiveSlide(initialColor, newColor) {
  PowerPoint.run(async (context) => {
    try {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items"); // Load the selected slide(s)
      await context.sync();

      if (selectedSlides.items.length === 0) {
        alert("No active slide selected.");
        return;
      }

      const slide = selectedSlides.items[0]; // Use the first selected slide
      console.log(`Applying color change to active slide: ${slide.id}`);

      // Load shapes (but exclude textFrame for now to prevent errors)
      const shapes = slide.shapes.load("items/name,items/fill,items/lineFormat");
      await context.sync();

      for (let shape of shapes.items) {
        console.log(`Checking shape: "${shape.name}"`);

        // Update fill color
        if (shape.fill) {
          shape.fill.load("foregroundColor");
          await context.sync();
          if (rgbOrHexEquals(shape.fill.foregroundColor, initialColor)) {
            shape.fill.setSolidColor(newColor);
            console.log(`Updated fill color for shape: ${shape.name}`);
          }
        }

        // Update line color
        if (shape.lineFormat) {
          shape.lineFormat.load("color");
          await context.sync();
          if (rgbOrHexEquals(shape.lineFormat.color, initialColor)) {
            shape.lineFormat.color = newColor;
            console.log(`Updated line color for shape: ${shape.name}`);
          }
        }

        // Check if shape supports text before attempting to load textFrame
        if (shape.textFrame) {
          try {
            shape.textFrame.load("textRange/font");
            await context.sync();
            if (rgbOrHexEquals(shape.textFrame.textRange.font.color, initialColor)) {
              shape.textFrame.textRange.font.color = newColor;
              console.log(`Updated text color for shape: ${shape.name}`);
            }
          } catch (textError) {
            console.warn(`Text color not updated for shape: ${shape.name}. Error: ${textError.message}`);
          }
        }
      }

      await context.sync(); // Sync changes
      console.log("Color change completed for active slide.");
    } catch (error) {
      console.error("Error changing colors on the active slide:", error);
    }
  });
}

// Flecha giratorias para seccion desplegable
function setupToggleWithArrow(buttonId, sectionId) {
  const $button = $(`#${buttonId}`);
  const $section = $(`#${sectionId}`);

  // Asegura que no haya salto visual al primer clic
  $section.css("display", "block");
  $button.addClass("open");

  // Cambia a slideToggle solo después de una pausa inicial
  setTimeout(() => {
    $section.css("display", ""); // quita el display: block inline para que slideToggle funcione
    $button.off("click").on("click", function() {
      $button.toggleClass("open");
      if ($button.hasClass("open")) {
        $section.slideDown(200);
      } else {
        $section.slideUp(200);
      }
    });
  }, 50); // da tiempo a que jQuery registre el estado inicial
}

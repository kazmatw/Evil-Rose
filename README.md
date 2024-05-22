# Tetris VBA Project

## Description

This project is a Tetris game developed in VBA (Visual Basic for Applications) for Microsoft Excel. It includes various modules that handle the game logic, UI updates, key handlers, timer events, and utilities.

## Preview
- ![image](https://github.com/kazmatw/VBA-TETRIS/assets/61039945/356d46dd-5848-463d-a512-2e8d0699766d)

## Project Structure

- **Microsoft Excel Objects**

  - `ThisWorkbook`: Contains code for initializing the game when the workbook is opened.
  - `Game (Game)`: The worksheet where the game is displayed.

- **Modules**
  - `ModGameLogic`: Contains the game logic, including block movement and collision detection.
  - `ModGlobals`: Stores global variables and constants used throughout the project.
  - `ModInitialization`: Handles the initialization of the game, including setting up the game field and initial values.
  - `ModKeyHandlers`: Contains subroutines for handling key presses.
  - `ModTimerEvents`: Manages the game timer and related events.
  - `ModUI`: Manages the user interface, including drawing the game field and updating statistics.
  - `ModUtilities`: Contains utility functions used in various parts of the project.
  - `ModExport`: Contains subroutines to automatically export and import all VBA modules for efficient development with Git.

## Setup

### Prerequisites

- Microsoft Excel
- Basic knowledge of VBA

### Installation

1. **Clone the Repository**

   ```sh
   git clone https://github.com/kazmatw/VBA-TETRIS.git
   cd VBA-TETRIS
   ```

2. **Open the Excel File**

   - Open the `tetrisProject.xlsm` file in Microsoft Excel.

3. **Enable Macros**

   - Ensure that macros are enabled in Excel to allow the VBA code to run.

4. **Import VBA Modules (Optional)**
   - If you need to import the modules manually, run the `ImportModules` subroutine in the `ModuleManagement` module.

## Usage

### Starting the Game

- The game initializes automatically when the workbook is opened.

### Controls

- **Down Arrow**: Move block down
- **Space**: Move block down
- **Left Arrow**: Move block left
- **Right Arrow**: Move block right
- **x**: Rotate block clockwise
- **c**: Rotate block counterclockwise

### Exporting and Importing Modules

- **Export Modules**: Run the `ExportModules` subroutine to export all VBA modules to the `ExportedModules` folder.
- **Import Modules**: Run the `ImportModules` subroutine to import all VBA modules from the `ExportedModules` folder.

## Collaboration

### Version Control with Git

- **Commit Changes**: After making changes to the VBA code, export the modules and commit them to the repository.

  ```sh
  git add ExportedModules/
  git commit -m "Describe your changes"
  git push
  ```

- **Pull Changes**: Before starting new work, pull the latest changes from the repository.
  ```sh
  git pull
  ```

### Handling Modules

- **Exporting Modules**: Use the `ExportModules` subroutine to export the VBA modules to the `ExportedModules` folder.
- **Importing Modules**: Use the `ImportModules` subroutine to import the VBA modules from the `ExportedModules` folder.

## Contributing

1. Fork the repository.
2. Create a new branch for your feature or bugfix.
   ```sh
   git checkout -b feature/your-feature
   ```
3. Make your changes and commit them.
   ```sh
   git commit -m "Describe your feature"
   ```
4. Push to your branch.
   ```sh
   git push origin feature/your-feature
   ```
5. Create a pull request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

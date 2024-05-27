# ColumnManager

A C# project demonstrating COM interop capabilities to interact with the Windows Shell to manage column information in Windows Explorer.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Examples](#examples)
- [Contributing](#contributing)
- [License](#license)

## Introduction

ColumnManager is a C# project that showcases how to use COM interop to interact with the Windows Shell for managing column information in Windows Explorer. This project includes examples of how to define and use COM interfaces, retrieve column information, and manipulate column settings.

## Features

- Interact with Windows Shell using COM interop
- Retrieve and manipulate column information in Windows Explorer
- Examples of handling common COM interop tasks

## Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download)
- Visual Studio or another C# development environment
- [Microsoft CsWin32](https://github.com/microsoft/CsWin32) - A source generator for generating P/Invoke and COM interop code.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/shea-c4/ColumnManager.git
    cd ColumnManager
    ```

2. Open the project in Visual Studio:
    - Open `ColumnManager.sln` in Visual Studio.

3. Add the `CsWin32` NuGet package:
    ```sh
    dotnet add package Microsoft.Windows.CsWin32
    ```

4. Create a `NativeMethods.txt` file in the project root directory. List the APIs you want to generate interop code for in this file. For example:
    ```
    GetWindowText
    GetWindowTextLength
    FindWindow
    ```

5. Add the following `ItemGroup` and `PropertyGroup` to your `.csproj` file to include the `NativeMethods.txt` file and enable `CsWin32` code generation:
    ```xml
    <ItemGroup>
      <PackageReference Include="Microsoft.Windows.CsWin32" Version="0.1.506-beta" />
    </ItemGroup>

    <PropertyGroup>
      <CsWin32GenerateFromNativeMethodsTxt>true</CsWin32GenerateFromNativeMethodsTxt>
    </PropertyGroup>

    <ItemGroup>
      <AdditionalFiles Include="NativeMethods.txt" />
    </ItemGroup>
    ```

## Usage

1. Build the project:
    ```sh
    dotnet build
    ```

2. Run the examples:
    ```sh
    dotnet run c:\windows --arch x64 --project ColumnManagerCLI
    ```

3. To use the COM interop functionality in your own project, reference the project or copy the relevant code files.


# Weather Dashboard Excel Add-In

## Overview

This project is a simple **Excel Add-In** designed for educational purposes to demonstrate the basic functionality of integrating external data into Excel. It allows users to fetch real-time weather data, such as temperature, humidity, and weather description, directly into an Excel worksheet by using the **OpenWeatherMap API**.

The add-in includes:
- A text box to input the city name.
- A button to trigger the weather data fetch.
- A table in the Excel worksheet that displays the fetched weather information.

## Features
- Fetches real-time weather data from OpenWeatherMap API.
- Displays city name, weather description, temperature, and humidity in an organized table in Excel.
- Simple interface with a text box for city input and a button to fetch data.

## Installation

1. Install **Visual Studio** (Community version recommended).
2. Clone this repository to your local machine.
3. Open the solution in **Visual Studio**.
4. Build and run the project.
5. Open Excel, and you should see the add-in available in the ribbon.
6. Enter a city name and click the button to fetch weather data.
### IMPORTANT: API Key Required

To make the project work, you **need to set up your own OpenWeatherMap API key**.
## Prerequisites

- **Visual Studio** (with Office/SharePoint Development workload).
- **OpenWeatherMap API key**. You can sign up [here](https://openweathermap.org/api) to get an API key.

## How It Works

1. The user inputs a city name into the text box.
2. The user clicks the "Fetch Weather" button.
3. The add-in sends a request to the **OpenWeatherMap API**.
4. The weather data is parsed and displayed in an Excel worksheet.

## Example

| City       | São Paulo      |
|------------|----------------|
| Description| Clear Sky      |
| Temperature| 28°C           |
| Humidity   | 65%            |


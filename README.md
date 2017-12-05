# README #

This is the data-mining repository for the students team:

* [Ankur Bahre](mailto:ankur.bahre@st.ovgu.de)

## Goal ##

Create a model to predict the temperatures from the output_section_temperature file.

## Background ##

* Heater with 16 sections and one common heating room to heat metall blocks
* each section has an own heating unit including an own temperature measurement
* based on air flow and heat exchange, everything is connected
* Further environment conditions:
    - heater walls will store temperature as well
    - metall blocks with different temperatures

## Zone influence

* Material flows from right to left
* Temperature flows from bottom to top and from left to right (e.g. Zone 8 is influenced by zone 12 and influences zones 6 and 4)

## Data set ##

You can use the following files to create your model:

### input_section_heating_energy ###
* 17 columns
    * 1. timestamp in s
    * 2.-17. energy of 16 sections in MW

### input_material_temperature ###
* 17 columns
    * 1. timestamp in s
    * 2-17. Metall block temperature in °C
* in between the given timestamp the temperature could be assumed as constant
* these values are computed by a model and could be not exact

### input_air_temperature ###
* 2 columns
    1. timestamp in s
    2. blow in air temperature in °C

### input_wall_temperature ###
* 17 columns
    * 1. timestamp in s
    * 2.-17. temperature of the walls in each seaction in °C
* these are the walls from the heater. Probably they should store some temperature.
* not so important

### output_section_temperature
* 17 columns
    * 1. timestamp in s
    * 2.-17. temperature in each seaction in °C
* this is your target. Your model should be able to predict this.

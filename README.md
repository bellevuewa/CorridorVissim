# CorridorVissim - Signal Control and Coordination Scripts

A Python-based toolkit for **traffic signal control** in VISSIM traffic simulation environments.

## Overview

CorridorVissim provides automated scripts to manage traffic signal timing and coordination within VISSIM simulation models. This project enables dynamic signal control, volume management, and coordinated traffic operations across corridor networks.

## Key Features

- **Signal Control**: Automated traffic signal timing management and optimization
- **Signal Coordination**: Coordinated signal timing across multiple intersections
- **Volume Management**: Import and adjust traffic volumes for realistic scenario testing
- **VISSIM Integration**: Direct integration with VISSIM simulation software
- **Flexible Configuration**: Configurable parameters for different scenario requirements

## Project Structure

```
vissim/
├── control_signal.py          # Signal control and coordination logic
├── run_vissim.py               # Main VISSIM simulation runner
├── adjust_volume.py            # Traffic volume adjustment utilities
├── import_volume.py            # Volume data import functionality
├── config_vissim.py            # VISSIM-specific configuration
├── utils.py                    # Common utility functions
└── logger.conf                 # Logging configuration
```

## Getting Started

### Prerequisites
- Python 3.x
- VISSIM (traffic simulation software)

### Setup
1. Clone the repository
2. Configure settings in `config.py` and `vissim/config_vissim.py`
3. Run simulations using the provided scripts

## Configuration

- **config.py**: Main project configuration
- **vissim/config_vissim.py**: VISSIM-specific parameters

## Usage

### Running VISSIM Simulations
Use `run_vissim.py` to execute traffic simulations with configured signal timings and parameters:
```python
python vissim/run_vissim.py
```

### Importing and Configuring Volumes
Use `import_volume.py` to import traffic volume data and `adjust_volume.py` to configure volumes for different scenarios:

```python
python vissim/import_volume.py    # Import volume data
```

For detailed usage instructions, refer to individual script docstrings and configuration files.


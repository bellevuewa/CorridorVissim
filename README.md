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
├── control_signal.py                # Signal control and coordination logic
├── run_vissim.py                    # Main VISSIM simulation runner (sequential)
├── run_vissim_threaded.py           # Threaded version using ThreadPoolExecutor
├── run_vissim_queueworkers.py       # Threaded version using worker queue pattern
├── adjust_volume.py                 # Traffic volume adjustment utilities
├── import_volume.py                 # Volume data import functionality
├── config_vissim.py                 # VISSIM-specific configuration
├── utils.py                         # Common utility functions
└── logger.conf                      
```

## Getting Started

### Prerequisites
- Python 3.9+
- VISSIM (traffic simulation software)

### Setup
1. Clone the repository
2. Configure settings in `config.py` and `vissim/config_vissim.py`
3. Run simulations using the provided scripts

## Configuration

- **config.py**: Main project configuration
- **vissim/config_vissim.py**: VISSIM-specific parameters

## Usage

### Running VISSIM Simulations - Choose Your Implementation

#### Sequential Version (Baseline)
Standard implementation with coordination running on the main thread:
```python
python vissim/run_vissim.py
```

#### ThreadPoolExecutor Version
Parallel coordination logic using thread pools. Adjust `max_workers` parameter in the script for your CPU configuration:
- `max_workers=2` or
- `max_workers=3`
and so on...

```python
python vissim/run_vissim_threaded.py
```

#### Queue Workers Version
Persistent worker threads using producer-consumer pattern. Adjust `num_workers` similarly to ThreadPoolExecutor version:

```python
python vissim/run_vissim_queueworkers.py
```

### Importing and Configuring Volumes
Use `import_volume.py` to import traffic volume data and `adjust_volume.py` to configure volumes for different scenarios:

```python
python vissim/import_volume.py    # Import volume data
```

For detailed usage instructions, refer to individual script docstrings and configuration files.
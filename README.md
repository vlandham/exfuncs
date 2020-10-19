# exfuncs

Wrapper around openpyxl that provides a set of composable functions for cleaning Excel files.

The hope is to provide a set of common excel transform functions that anyone can use to get their excel file into a more computer readable format.

## Documentation

_TODO_

## Local Development

### Dependency Installation

Change to a package directory folder and install the dependencies for it:

```bash
pipenv sync --dev  # the --dev include packages used for development
```

This creates a folder .venv so you can activate the virtual environment.

Then activate this virtual environment using:

```bash
pipenv shell # use "exit" to deactivate this environment
```

You only need to do pipenv sync --dev when Pipfile.lock has changed. Otherwise, after .venv
exists you can simply activate the virtual environment when
working on the project.

The `exfuncs` package is included in its own `Pipfile`, so installing the Pipenv dependencies in turn installs `exfuncs` as well.

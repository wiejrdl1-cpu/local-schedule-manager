"""Keep tkinter discoverable when Tcl cannot initialize during analysis."""


def pre_find_module_path(hook_api):
    # The build environment can live below a Korean Windows user path.  Tcl's
    # analysis probe may fail to initialize there even though tkinter itself is
    # installed and usable in the frozen application.
    return None

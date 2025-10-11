# pyscript/core/__init__.py

# --- Dynamic import shim for frozen builds ---
try:
    __all__
except NameError:
    __all__ = []

def __getattr__(name):
    if name == "core_voice":
        import importlib
        return importlib.import_module(__name__ + ".core_voice")
    raise AttributeError(name)

if "core_voice" not in __all__:
    __all__.append("core_voice")
# --- End shim ---

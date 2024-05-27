# Python SAPI ISpTTSEngine Implementation Proof-of-Concept

This project provides a proof-of-concept implementation of SAPI Text-to-Speech
Engine in Python. It can be used as a starting point for developing a real SAPI
compliant speech synthesizer.

## Open Developer Command Prompt and run MIDL Compiler

```cmd
midl.exe pysapi.idl
```

## Generate COM types

```cmd
python generate_com_types.py
```

## Register SAPI voice (must run as Administrator)

```cmd
python sapi_tts_engine.py /regserver
```

## Seatback and enjoy :-)

```cmd
python speak.py "Python SAPI PoC Voice"
```

## Useful references

- [COM servers with comtypes](https://pythonhosted.org/comtypes/server.html)
- [Speech API Overview (SAPI 5.4)](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ee125077(v=vs.85))
- [tts - A lightweight python TTS wrapper](https://github.com/DeepHorizons/tts)

# Object Repository

Put your shared object repositories (`.tsr`) here.

Create one in UFT: **Resources > Object Repository Manager > File > New**, add your objects, then **Save As** into this folder (for example `Application.tsr`).

By default every `*.tsr` in this folder is loaded at runtime. To load specific files only, set `OBJECT_REPOSITORY` in `.env`, for example:

```
OBJECT_REPOSITORY=ObjectRepository\Login.tsr;ObjectRepository\Home.tsr
```

No repository is required to start: tests can use descriptive programming (see `FunctionLibrary/AppFunctions.vbs`).

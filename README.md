# AV2 Codec ISO Media File Format Binding (AV2-ISOBMFF)

This repository hosts the specification work for encapsulating AV2 bitstreams in ISO Base Media File Format (ISOBMFF).

## Related repositories

- [av2-tools](https://github.com/AOMediaCodec/av2-tools) - reference software implementing this specification (AV2 OBU parsing/analysis and MP4/ISOBMFF muxing & demuxing).
- [av2-interop-private](https://github.com/AOMediaCodec/av2-interop-private) - AV2 interop stream recipes and catalog used to exercise the toolchain against this spec.

## Specification source

The specification is written using a special syntax (mixing markup and markdown) to enable generation of cross-references, syntax highlighting, ...
The file using this syntax is [index.bs](./index.bs).

[index.bs](./index.bs) is processed by [specbuild](./tools/specbuild/), a Bikeshed-based build pipeline, when content is pushed onto the `main` branch or when Pull Requests are made.

## Building locally

Requirements: Python 3.10+ and Node.js 20+.

It is recommended to use a dedicated Python environment:

```shell
python3 -m venv venv
source venv/bin/activate
```

Install Python and Node dependencies:

```shell
pip install -r tools/specbuild/requirements.txt
(cd tools/specbuild && npm install)
```

Compile the spec:

```shell
python tools/specbuild/compile.py --profile publication --toc-leaders table --check-sdl-syntax
```

The compiled output is written to `av2-isobmff_Spec/`. See [tools/specbuild/README.md](./tools/specbuild/README.md) for additional build options (PDF, multipage, diff against another revision, etc.).

# Node.js Demo

## Regular Node Demo

### Regular Usage

Generate a simple presentation.

```bash
node demo.js
```

Generate a presentation containing all demo objects (equivalent to the browser demo).

```bash
node demo.js All
```

Generate a presentation containing selected demo objects (e.g. 'Table', 'Text').
See `../common/demos.js` for the full list of tests.

```bash
node demo.js Text
```

## Stream Demo

The `demo_stream.js` file requires the `express` package to demonstrate streaming.

### Stream Usage

```bash
node demo_stream.js
```

Then open `http://localhost:3000/` in a local web browser to download the streamed file.

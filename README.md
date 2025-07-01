# Outlook Plugin

This project uses a configurable `BASE_URL` to define the domain for the add-in.
The repository includes a `.env` file preconfigured with the current domain.
If you need to change the domain, edit `.env` before building. The default is:

```
BASE_URL=https://9a08-92-64-101-76.ngrok-free.app
```

Run the build with:

```bash
npm run build
```

`BASE_URL` will be injected into `manifest.xml` and the code at build time.


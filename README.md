# AutoLog

A maintenance log for every car you own. Plain HTML, CSS and JavaScript, with no build step and no server.

## What it does

- Car details: make, model, trim, plate, engine number, VIN (typed or read from a photo), tires, oil, filters, wipers and lights.
- Maintenance log: date, work done, location, mileage and cost, with a running total. Sort by any column and filter by text, location, date range or oil changes.
- Oil change reminder every 6 months, counted from the last oil change you logged. Browser notification when you open the site, plus a repeating calendar event you can download.
- Backup to an Excel spreadsheet and restore from it.
- Links to parts stores near you.
- Profiles, so several people can keep separate cars and logs.

## Where the data lives

Everything is saved in the visitor's own browser (localStorage). Nothing is uploaded, and each visitor to the site only ever sees their own records. To move records to another device, download a backup there and restore it here.

## Publish on GitHub Pages

1. Create a repository on GitHub and push these files to the `main` branch.
2. In the repository, open Settings, then Pages.
3. Under "Build and deployment", choose "Deploy from a branch", pick `main` and `/ (root)`, and save.
4. The site appears at `https://<your-username>.github.io/<repository-name>/` after a minute or two.

## Run it locally

Open `index.html` in a browser. Notifications need the site to be served over `https` or from `localhost`, for example `npx serve`.

## Credits

Hero photo: [1969 Ford Mustang Mach 1 in Red and Black](https://commons.wikimedia.org/wiki/File:1969_Ford_Mustang_Mach_1_in_Red_and_Black_(53520101274).jpg) by crudmucosa, CC BY 2.0, resized.
Spreadsheet support by SheetJS and VIN photo reading by Tesseract.js, both loaded on demand from their CDNs.

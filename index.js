const docx = require('docx');
const fs = require('fs');
const glob = require('glob');

const { promisify } = require('util');

const globProm = promisify(glob);
const readFile = promisify(fs.readFile);
const writeFile = promisify(fs.writeFile);

const {
    Document,
    ImageRun,
    Paragraph,
    Packer,
    HorizontalPositionRelativeFrom,
    HorizontalPositionAlign,
    VerticalPositionAlign,
    VerticalPositionRelativeFrom
} = docx;

async function start() {

  console.log('Loading images...');
  const imagesPaths = await globProm('./images/*.*');
  console.log('Images loaded');

  console.log('Sorting images...');
  imagesPaths.sort((f1, f2) => {
    const fileName1 = f1.substring(f1.lastIndexOf('/') + 1, f1.lastIndexOf('.'));
    const fileName2 = f2.substring(f2.lastIndexOf('/') + 1, f2.lastIndexOf('.'));
    if (Number.isNaN(parseInt(fileName1)) ||  Number.isNaN(parseInt(fileName2))) {
      const a = fileName1.replace(/\s/gi,'');
      const b = fileName2.replace(/\s/gi,'');
      return a > b ? 1 : -1;
    } else {
      return parseInt(fileName1) - parseInt(fileName2);
    }
  });
  console.log('Images sorted');

  const imagesProm = imagesPaths.map(async imagePath => {
    console.log('Adding image ', imagePath);
    const imageRaw = await readFile(imagePath);
    console.log('Added image ', imagePath);

    return {
      children: [
        new Paragraph({
          children: [
            new ImageRun({
              data: imageRaw,
              transformation: {
                width: 800,
                height: 1100,
              },
              floating: {
                horizontalPosition: {
                    relative: HorizontalPositionRelativeFrom.PAGE,
                    align: HorizontalPositionAlign.CENTER
                },
                verticalPosition: {
                    relative: VerticalPositionRelativeFrom.PAGE,
                    align: VerticalPositionAlign.CENTER,
                },
            },
            })
          ],
        })
      ]
    };
  }); 
  
  const images = await Promise.all(imagesProm);

  const doc = new Document({
    sections: images
  });

  console.log('Generating document...');
  const buffer = await Packer.toBuffer(doc)
  console.log('Document generated');

  console.log('Saving document...');
  await writeFile(`File ${new Date().getTime()}.docx`, buffer);
  console.log('Document saved. Done!');
}

start();
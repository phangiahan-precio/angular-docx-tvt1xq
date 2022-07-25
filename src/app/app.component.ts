import { Component } from '@angular/core';
import { Packer } from 'docx';
import { saveAs } from 'file-saver/FileSaver';

import {
  experiences,
  education,
  skills,
  achievements,
  languages,
  commitments,
  employers,
} from './cv-data';
import { DocumentCreator } from './cv-generator';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  name = 'Angular';

  public download(): void {
    const documentCreator = new DocumentCreator();
    const doc = documentCreator.create([
      experiences,
      education,
      skills,
      achievements,
      languages,
      commitments,
      employers,
    ]);

    Packer.toBlob(doc).then((blob) => {
      console.log(blob);
      saveAs(blob, 'example.docx');
      console.log('Document created successfully');
    });
  }
}

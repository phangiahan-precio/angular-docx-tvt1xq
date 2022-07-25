import {
  AlignmentType,
  convertInchesToTwip,
  Document,
  HeadingLevel,
  ImageRun,
  Paragraph,
  TabStopPosition,
  TabStopType,
  TextRun,
  UnderlineType,
  PageBreak,
  Table,
  TableRow,
  TableCell,
  BorderStyle,
  VerticalAlign,
} from 'docx';

const imageBase64Data = `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`;

export class DocumentCreator {
  public create([
    experiences,
    educations,
    skills,
    achivements,
    languages,
    commitments,
    employers,
  ]): Document {
    const document = new Document({
      styles: {
        default: {
          title: {
            run: {
              size: 48,
              font: 'Calibri (Body)',
              color: 'e57e14',
            },
          },

          heading1: {
            run: {
              size: 28,
              color: '898d8f',
              font: 'Calibri (Body)',
            },
            paragraph: {
              spacing: {
                after: 120,
              },
            },
          },
          heading2: {
            run: {
              size: 20,
              bold: true,
              underline: {
                type: UnderlineType.DOUBLE,
                color: '898d8f',
              },
            },
            paragraph: {
              spacing: {
                before: 240,
                after: 120,
              },
            },
          },
          listParagraph: {
            run: {
              color: '#FF0000',
            },
          },
        },
        paragraphStyles: [
          {
            id: 'skillsPart',
            name: 'SkillsPart',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              color: '434343',
              italics: false,
              underline: { color: '434343', type: undefined },
              font: 'Calibri (Body)',
              size: 20,
            },
          },
          {
            id: 'skillsPartContent',
            name: 'SkillsPartContent',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              color: '8c9092',
              italics: true,
              font: 'Calibri (Body)',
              size: 22,
            },
            paragraph: {
              spacing: {
                line: 276,
                before: 20 * 72 * 0.1,
                after: 20 * 72 * 0.05,
              },
            },
          },
          {
            id: 'normalText',
            name: 'NormalText',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              color: '434343',
              font: 'Calibri (Body)',
              size: 22,
            },
            paragraph: {
              spacing: {
                line: 276,
                before: 20 * 72 * 0.1,
                after: 20 * 72 * 0.05,
              },
            },
          },
          {
            id: 'tableStyle',
            name: 'TableStyle',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              border: { style: BorderStyle.NONE },
            },
          },
        ],
      },

      sections: [
        {
          children: [
            //page 1
            this.createHeadingTitle('Dolan Miu', 'Fullstack Developer'),
            this.createInfo(
              'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
            ),
            new Table({
              borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                insideVertical: { style: BorderStyle.NONE },
              },
              style: 'tableStyle',
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      borders: {
                        top: { style: BorderStyle.NONE },
                        bottom: { style: BorderStyle.NONE },
                        left: { style: BorderStyle.NONE },
                        right: { style: BorderStyle.NONE },
                      },
                      children: [
                        new Paragraph({
                          children: [
                            new ImageRun({
                              data: Buffer.from(imageBase64Data, 'base64'),
                              transformation: {
                                width: 100,
                                height: 100,
                              },
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      verticalAlign: VerticalAlign.BOTTOM,
                      children: [
                        new Table({
                          borders: {
                            top: { style: BorderStyle.NONE },
                            bottom: { style: BorderStyle.NONE },
                            left: { style: BorderStyle.NONE },
                            right: { style: BorderStyle.NONE },
                          },
                          style: 'tableStyle',
                          rows: [
                            new TableRow({
                              children: [
                                new TableCell({
                                  children: [
                                    new Paragraph(
                                      'Hello1111111111111111111111111111111111111111111111'
                                    ),
                                    new Paragraph('Hello'),
                                  ],
                                }),
                              ],
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            new Paragraph({
              border: {
                bottom: {
                  color: '#fd7e14',
                  space: 1,
                  style: BorderStyle.SINGLE,
                  size: 10,
                },
              },
            }),

            //page 2
            this.createHeading('Education'),
            ...educations
              .map((education) => {
                const arr: Paragraph[] = [];
                arr.push(
                  this.createInstitutionHeader(
                    education.schoolName,
                    `${education.startDate.year} - ${education.endDate.year}`
                  )
                );
                return arr;
              })
              .reduce((prev, curr) => prev.concat(curr), []),
            this.createHeading('Project & Assignments'),
            ...educations
              .map((education) => {
                const arr: Paragraph[] = [];

                arr.push(
                  this.createWorkExperiencesHeader(
                    `${education.startDate.year} - ${education.endDate.year}`,
                    education.schoolName
                  ),
                  new Paragraph({
                    text: 'NGO Online is a cloud-based program, project and grant management IT solution specifically designed for international humanitarian and development NGOs. Powerful enough to meet your information management needs but lightweight enough to use in low-bandwidth areas, NGO Online supports everything from grant management through to operations in the field',
                    style: 'normalText',
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                      left: 1450,
                    },
                  }),
                  this.createWorkExperienceSkillList(skills)
                );

                return arr;
              })
              .reduce((prev, curr) => prev.concat(curr), []),
            this.createHeading('Skill'),
            this.createSubHeading('LEVEL 5 - EXPERT'),
            this.createSkillList(skills),
            this.createSubHeading('LEVEL 4 - ADVANCED'),
            this.createSkillList(skills),
            this.createSubHeading('LEVEL 3 - HIGH COMPETENCE'),
            this.createSkillList(skills),
            this.createSubHeading('LEVEL 2 - INTERMEDIATE'),
            this.createSkillList(skills),
            this.createSubHeading('LEVEL 1 - NOVICE'),
            this.createSkillList(skills),

            this.createHeading('Languages'),
            ...languages
              .map((l) => {
                const arr: Paragraph[] = [];
                arr.push(
                  this.createInstitutionHeader(
                    l.name,
                    this.getLevelStatus(l.level)
                  )
                );
                return arr;
              })
              .reduce((prev, curr) => prev.concat(curr), []),

            this.createHeading('Commitments and publications'),
            ...commitments
              .map((l) => {
                const arr: Paragraph[] = [];
                arr.push(
                  this.createInstitutionHeader(
                    l.title,
                    this.createPositionDateText(
                      l.startDate,
                      l.endDate,
                      l.onGoing
                    )
                  )
                );
                return arr;
              })
              .reduce((prev, curr) => prev.concat(curr), []),
          ],
        },
      ],
    });
    return document;
  }

  public createInfo(text: string): Paragraph {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      indent: {
        right: 4000,
      },
      children: [
        new TextRun({
          text: text,
        }),
        new PageBreak(),
      ],
    });
  }

  public createHeading(text: string): Paragraph {
    return new Paragraph({
      text: text,
      heading: HeadingLevel.HEADING_1,
      thematicBreak: true,
    });
  }

  public createHeadingTitle(name: string, position: string): Paragraph {
    return new Paragraph({
      children: [
        new ImageRun({
          data: Buffer.from(imageBase64Data, 'base64'),
          transformation: {
            width: 100,
            height: 100,
          },
        }),
        new Paragraph({
          text: name,
          heading: HeadingLevel.TITLE,
          thematicBreak: true,

          children: [
            new Paragraph({
              text: position,
              heading: HeadingLevel.HEADING_4,
            }),
          ],
        }),
        // new Paragraph({
        //   text: position,
        //   heading: HeadingLevel.HEADING_4,
        //   indent: {
        //     left: 1500,
        //   },
        // }),
      ],
    });
  }

  public createSubHeading(text: string): Paragraph {
    return new Paragraph({
      text: text,
      heading: HeadingLevel.HEADING_2,
      style: 'skillsPart',
    });
  }

  public createInstitutionHeader(
    institutionName: string,
    dateText: string
  ): Paragraph {
    return new Paragraph({
      style: 'normalText',
      tabStops: [
        {
          type: TabStopType.RIGHT,
          position: TabStopPosition.MAX,
        },
      ],
      children: [
        new TextRun({
          text: institutionName,
        }),
        new TextRun({
          text: `\t${dateText}`,
        }),
      ],
    });
  }

  public createWorkExperiencesHeader(
    institutionName: string,
    dateText: string
  ): Paragraph {
    return new Paragraph({
      style: 'normalText',
      tabStops: [
        {
          type: TabStopType.RIGHT,
          position: 50,
        },
      ],
      children: [
        new TextRun({
          text: institutionName,
        }),
        new TextRun({
          text: `\t${dateText}`,
        }),
      ],
    });
  }

  public createRoleText(roleText: string): Paragraph {
    return new Paragraph({
      children: [
        new TextRun({
          text: roleText,
          italics: true,
        }),
      ],
    });
  }

  public createBullet(text: string): Paragraph {
    return new Paragraph({
      text: text,
      bullet: {
        level: 0,
      },
    });
  }

  public createSkillList(skills: any[]): Paragraph {
    return new Paragraph({
      style: 'skillsPartContent',
      children: [
        new TextRun(skills.map((skill) => skill.name).join(', ') + '.'),
      ],
    });
  }

  public createWorkExperienceSkillList(skills: any[]): Paragraph {
    return new Paragraph({
      style: 'skillsPartContent',
      children: [
        new TextRun(skills.map((skill) => skill.name).join(', ') + '.'),
      ],
      indent: {
        left: 1450,
      },
    });
  }

  getLevelStatus(level: any) {
    switch (level) {
      case 1:
        return 'No Proficiency';
      case 2:
        return 'Elementary Proficiency';
      case 3:
        return 'Professional Working Proficiency';
      case 4:
        return 'Full Professional Proficiency';
      case 5:
        return 'Native Or Bilingual Proficiency';
      default:
        return '';
    }
  }

  // tslint:disable-next-line:no-any
  public createAchivementsList(achivements: any[]): Paragraph[] {
    return achivements.map(
      (achievement) =>
        new Paragraph({
          text: achievement.name,
          bullet: {
            level: 0,
          },
        })
    );
  }

  public createInterests(interests: string): Paragraph {
    return new Paragraph({
      children: [new TextRun(interests)],
    });
  }

  public splitParagraphIntoBullets(text: string): string[] {
    return text.split('\n\n');
  }

  // tslint:disable-next-line:no-any
  public createPositionDateText(
    startDate: any,
    endDate: any,
    isCurrent: boolean
  ): string {
    const startDateText =
      this.getMonthFromInt(startDate.month) + ' ' + startDate.year;
    const endDateText = isCurrent
      ? 'Ongoing'
      : `${this.getMonthFromInt(endDate.month)} ${endDate.year}`;

    return `${startDateText} - ${endDateText}`;
  }

  public getMonthFromInt(value: number): string {
    switch (value) {
      case 1:
        return 'Jan';
      case 2:
        return 'Feb';
      case 3:
        return 'Mar';
      case 4:
        return 'Apr';
      case 5:
        return 'May';
      case 6:
        return 'Jun';
      case 7:
        return 'Jul';
      case 8:
        return 'Aug';
      case 9:
        return 'Sept';
      case 10:
        return 'Oct';
      case 11:
        return 'Nov';
      case 12:
        return 'Dec';
      default:
        return 'N/A';
    }
  }
}

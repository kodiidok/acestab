// use the "i" flag for case-insensitivity
// const regex_degree_title = /honours|hons|special|B\.?Sc\.?|BVSc\.?|B\.?V\.?S\.?|V\.?M\.?S\.?|MBBS|BDA|B\.?A\.?|M\.?Sc\.?|Master\s?|PhD|M\.?Phil\.?/i; 
// const regex_degree_subject = /Computer Science|Botany|Zoology|Chemistry/i;

const degrees = {};

function doGet(e) {
  let htmlOutput = '';
  if (!e.parameter.page) {
    // When no specific page requested, return "home page" Ex : ?page=hod
    htmlOutput = HtmlService.createTemplateFromFile('index');
    return htmlOutput
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // else, use page parameter to pick an html file from the script

  htmlOutput = HtmlService.createTemplateFromFile(e.parameter.page);
  return htmlOutput
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function main () {
  try {
    const ss = SpreadsheetApp.openById('sheet_id');
    
    // const bdegree_sheet_name = 'basicDegreeTitles';
    // const pgdegree_sheet_name = 'pgDegreeTitles';
    // const subjectAreas_sheet_name = 'subjectAreas';

    const bdegree_sheet_name = 'basicDegreeTitles';
    const pgdegree_sheet_name = 'postgradDegreeTitles';
    const subjectAreas_sheet_name = 'subjectAreas';

    const bdegree_ws = ss.getSheetByName(bdegree_sheet_name);
    const pgdegree_ws = ss.getSheetByName(pgdegree_sheet_name);
    const subjects_ws = ss.getSheetByName(subjectAreas_sheet_name);
    
    degrees.bdegree_titles = bdegree_ws.getDataRange().getValues();
    degrees.pgdegree_titles = pgdegree_ws.getDataRange().getValues();
    degrees.subjects = subjects_ws.getDataRange().getValues();

    saveRequest();
  } catch (e) {
    console.error(`Error occured while opening the sheet. error : `, e);
    throw new Error(`Error occured while opening the sheet : `, e);
  }
}

/** This method will take an array of DEGREE_TITLES as a parameter. By iterating through each DEGREE_TITLE in the array, it will seperate each character and add `\\.?` between each character. Then, the new strings are pushed to a new array, which will contain the elements needed for the `regex`. Finally, the regular expression for the DEGREE_TITLES is returned. */
function genRegex (retdata, type) {
  let regex_degree = '';
  const newdata = [];
  if (retdata !== 'none') {
    const data = retdata.map(([element]) => element); // const data = ['BSc', 'BVSc', 'BVS', 'VMS', 'BA', 'MBBS'];
    data.forEach((item) => {
      let temp = '\\.?';
      for (i=0; i<item.length; i+=1) {
        temp += item[i];
        temp += '\\.?';
      }
      newdata.push(temp);
    });
    if (type === 'BASIC_DEGREE') { // basic degree
      regex_degree = new RegExp(`honours|hons|special|${newdata.join('|')}`, 'i');
    } else if (type === 'POSTGRAD_DEGREE') { // postgrad degree
      regex_degree = new RegExp(`${newdata.join('|')}`, 'i');
    } else if (type === 'SUBJECT') { // subject
      regex_degree = new RegExp(`${newdata.join('|')}`, 'i');
    }
  } else {
    if (type === 'CLASS') { // class
      regex_degree = new RegExp(`1|first|2|second|upper|lower`, 'i');
      // return /1|first|2|second|upper|lower/i
    }
  }
  // Logger.log(regex_degree.toString());
  return regex_degree;
}

/** loads data from database */
function loadsheet () {
  try {
    const ss = SpreadsheetApp.openById('sheet_id');
    const sheetName = 'mainsheet';
    const ws = ss.getSheetByName(sheetName);
    const data = ws.getDataRange().getValues();
    return data;
  } catch (e) {
    console.error(`Error occured while opening the sheet. When trying to open the sheet :${sheetName}error : `, e);
    throw new Error(`Error occured while opening the sheet : ${sheetName}`, e);
  }
}

/** loads the saving database and returns the mainsheet */
function savesheet () {
  try {
    const ss = SpreadsheetApp.openById('sheet_id');
    const sheetName = 'mainsheet';
    const ws = ss.getSheetByName(sheetName);
    return ws;
  } catch (e) {
    console.error(`Error occured while opening the sheet. When trying to open the sheet :${sheetName}error : `, e);
    throw new Error(`Error occured while opening the sheet : ${sheetName}`, e);
  }
}

/** duplicates the original sheet to prevent data loss when performing the shortlisting process */
function copysheet (data) {
  const db = savesheet();
  loadsheet().forEach((item) => {
    db.appendRow(item);
  });
}

/** reads each record from the database, performs the shortlisting process, and then saves the `short list status` in a new column */
function saveRequest () {
  try {
    const db = savesheet();
    const retdata = db.getDataRange().getValues();
    const data = retdata.slice(1);
    // const saveIndex = db.getLastColumn() + 1;
    const saveIndex = db.getLastColumn();
    const lastrow = db.getLastRow();
     
    let i = 2;
    data.forEach((item) => {
      const status = shortlist(item);
      if (status) {
        db.getRange(i, saveIndex).setValue(status);
        // Logger.log(status);
      }
      i++;
    });
  
    // data.forEach((item) => {
      // Logger.log(shortlist(item).vacancy);
      // Logger.log(`${shortlist(item).bdegree.name}\t${regex_degree_title.test(shortlist(item).bdegree.name)}`);
      // Logger.log(`${shortlist(item).bdegree.name}\t${genRegex(degrees.bdegree_titles, 1).test(shortlist(item).bdegree.name)}`);
      /**
        let str = '';
        shortlist(item).pgdegree.forEach((element) => {
          str += `${genRegex(degrees.pgdegree_titles, 2).test(element.name)}\n`;
        });
        Logger.log(str);
      */
      // Logger.log(shortlist(item).pgdegree);
      // Logger.log(shortlist(item).declaration);
      // Logger.log(shortlist(item).presentOccupation);
      // Logger.log(shortlist(item).previousOccupation);
    // });

  } catch (error) {
    console.error('Error occurred while saving data', error);
    throw new Error(`Error occurred while saving data`, error);
  }
}

/** short listing process */
function shortlist (data) {
  const application = data;

  /** 2. Vacancy ID    3. Post   4. Faculty   5. Department */
  const vacancy = {};
  vacancy.id = application[2];
  vacancy.post = application[3];
  vacancy.faculty = application[4];
  vacancy.department = application[5];
  
  /** 29. Basic Degree	  30. BD Country    31. BD University	  32. BD from	  33. BD to	  34. BD class	  35. BD gpa */
  const bdegree = {};
  bdegree.name = application[29];
  bdegree.country = application[30];
  bdegree.university = application[31];
  bdegree.bdfrom = application[32];
  bdegree.bdto = application[33];
  bdegree.bdclass = application[34];
  bdegree.gpa = application[35];

  /** 36. Postgrad Degree */
  // const pgdegree = JSON.parse(application[36]);
  const pgdegree = JSON.parse(application[36]).map((item) => {
    return {
      name: item[0],
      university: item[1],
      country: item[2],
      pgdfrom: item[3],
      pgdto: item[4],
      pgdclass: item[5],
      pgdgpa: item[6],
      pgdmethod: item[7]
    }
  });

  /** 41. Commendations	  42. CPDetails    31. Vacation Post Notice	  58. Bond Violation	  59. Bond Value	  60. Bond Institute */
  const declaration = {};
  declaration.CommendedOrPunished = application[41];
  declaration.CPDetails = application[42];
  declaration.VacationPostNotice  = application[43];
  declaration.BondViolations = application[58];
  declaration.BondValue = application[59];
  declaration.BondInstitute = application[60];

  /**  45. PRES from    46. PRES designation    47. PRES department   48. PRES salary */
  const presentOccupation = {};
  presentOccupation.presFrom = application[45];
  presentOccupation.presDesignation = application[46];
  presentOccupation.presDepartment = application[47];
  presentOccupation.presSalary = application[48];

  /**  57. Previous Occupation */
  const previousOccupation = application[57];

  /** vacancy, basic, postgrad, declaration, currentEmployement, previousEmployement */
  const status = checkQualifications(vacancy, bdegree, pgdegree, declaration, presentOccupation, previousOccupation);
  return status;
  // return { vacancy, bdegree, pgdegree, declaration, presentOccupation, previousOccupation };
}

/** calculates the duration of `years of experience` */
function calculateDurationInYears(startDate, endDate) {
  let end = '';
  const start = new Date(startDate);
  if (endDate === undefined || endDate === null) {
    end = new Date();
  } else {
    end = new Date(endDate);
  }
  const diffInTime = Math.abs(end.getTime() - start.getTime());
  const diffInMonths = Math.round(diffInTime / (1000 * 3600 * 24 * 30));
  const result = Math.ceil(diffInMonths / 12);
  if (isNaN(result)) {
    return 'Invalid Date Format!';
  }
  return result;
}

/** check qualifications and decide whether to `short list` or make the application a `pending` application accoring to a given set of parameters */
function checkQualifications (vacancy, bdegree, pgdegree, declaration, presentOccupation, previousOccupation) {
  
  let result = 'pending';
  
  if (bdegree !== null) {
    const bdtitle = genRegex(degrees.bdegree_titles, 'BASIC_DEGREE').test(bdegree.name);
    const bdsubject = genRegex(degrees.subjects, 'SUBJECT').test(bdegree.name);
    const bdclass = genRegex('none', 'CLASS').test(bdegree.bdclass);
    if (bdtitle && bdsubject) {
      if (bdclass) {
        /*
          If DEGREE_TITLE == Given title (title should add for Department by Ac.Est. staff)
            If  CLASS == 1st  /2nd upper HONS -> Shortlist
            Else If  CLASS == 2nd lower HONS -> Shortlist
              Else  Pending for manual shortlist
          Else if DEGREE_TITLE == Null -> Reject
          Else  Pending for manual shortlist
        */
        if (vacancy.post === 'Lecturer (Probationary)') { /** shortlist only if applicant has a basic degree with DEGREE_TITLE, SUBJECT and CLASS pre-defined by the academic est. */
          result = 'short listed';
        }
        const masters = [];
        pgdegree.forEach((item) => {
          const pgdtitle = genRegex(degrees.pgdegree_titles, 'POSTGRAD_DEGREE').test(item.name);
          const pgdsubject = genRegex(degrees.subjects, 'SUBJECT').test(item.name);
          if (pgdegree && pgdsubject) {
            const duration = calculateDurationInYears(item.pgdfrom, item.pgdto);
            /*
              If DEGREE_TITLE == Given title AND [(CLASS == 1st  /2nd upper HONS) OR ( CLASS == 2nd lower HONS )]
                  If PHD_TITLE == Given title -> Shortlist
                Else if Masters_TITLE == Given title AND Duration>2 years  shortlist          //(for Science masters)
                Else if Masters_TITLE == Given title AND Duration>=1.5 years   Pending for manual shortlist    
              Else if DEGREE_TITLE == Null -> Reject
              Else  Pending for manual shortlist
            */
            if (vacancy.post === 'Lecturer (Unconfirmed)') {
              if (item.name.charAt(0).toUpperCase() === 'P') {
                result = 'short listed';
              } else {
                if (item.name.charAt(0).toUpperCase() === 'M' && duration > 2) {
                  masters.push(true);
                } else {
                  masters.push(false);
                }
                masters.forEach((item) => {
                  if (!item) {
                    result = 'pending';
                  }
                  result = 'short listed';
                });
              }
            }
            /* 
              If DEGREE_TITLE == Given title (title should add for Department by Ac.Est. staff)  
                If  CLASS == 1st  /2nd upper HONS OR  CLASS == 2nd lower HONS
                  If PHD_TITLE == Given title   Shortlist
                  Else if (Masters_TITLE == Given title) AND {(Duration>2 years) OR (Duration>=1.5 years)}   Pending for manual shortlist  
            */
            if (vacancy.post === 'Senior Lecturer Grade II') {
              if (item.name.charAt(0).toUpperCase() === 'P') {
                result = 'short listed';
              } else {
                if (item.name.charAt(0).toUpperCase() === 'M' && duration > 1.5) {
                  result = 'pending';
                }
              }
            } 
            /** Pending for manual shortlist */
            if (vacancy.post === 'Senior Lecturer Grade I') {
              result = 'pending';
            }
          }
        });
      }
    }
  } else { /** reject applicant if there are no basic degrees */
    result = 'rejected';
  }

  return result;
}
  

import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as pdw from 'pdw';
import { Temporal } from 'temporal-polyfill';
// import * as YAML from 'yaml';


// export function importPreviousCSV(filepath: string): pdw.CompleteDataset{
//     console.log('loading...');
//         let returnData: pdw.CompleteishDataset = {}
//         XLSX.set_fs(fs);
//         let loadedWb = XLSX.readFile(filepath);
//         const shts = loadedWb.SheetNames;
//         const pdwRef = pdw.PDW.getInstance();
//         const sht = loadedWb.Sheets[shts[0]];
//         let elements = XLSX.utils.sheet_to_json(sht) as pdw.DefLike[];

//         let defs: any = [];
//         let pointDefs: any = [];
//         let entries: any = [];
//         let entryPoints: any = [];
//         let tagDefs: any = [];
//         let tags: any = [];

//         elements.forEach((element: any) => {
//             if (element['Row Type'] === 'Def') defs.push(buildElement(element));
//             if (element['Row Type'] === 'PointDef') pointDefs.push(buildElement(element));
//             if (element['Row Type'] === 'Entry') entries.push(buildElement(element));
//             if (element['Row Type'] === 'EntryPoint') entryPoints.push(buildElement(element));
//             if (element['Row Type'] === 'TagDef') tagDefs.push(buildElement(element));
//             if (element['Row Type'] === 'Tag') tags.push(buildElement(element));
//         })

//         // defs.forEach(def=>{
//         //     def._pts = pointDefs.filter(pd=>pd._did === def._did).map(p=>{
//         //         return {
//         //             _pid: p._pid,
//         //             _lbl: p._lbl,
//         //             _desc: p._desc,
//         //             _emoji: "🙋‍♂️",
//         //             _type: p._type,
//         //             _rollup: p._rollup,
//         //             _active: true,
//         //             _opts: []
//         //         }
//         //     });
//         //     def._emoji = "🙋‍♂️";
//         // })

//         let ewo = 0

//         entries.forEach(entry=>{
//             let points = entryPoints.filter(ep=> ep._eid === entry._eid)
//             if(points.length == 0){
//                 ewo = ewo + 1;
//                 //shoot.
//                 // console.log(points);
//             }
            
//         })

//         pdwRef.setDefs((<pdw.DefLike[]>defs))
//         pdwRef.setEntries((<pdw.EntryLike[]>entries))

//         return {
//             defs: defs,
//             entries: entries,
//             tags: []
//         }

//         function buildElement(elementData: any) {
//             let returnObj: any = {
//                 _uid: elementData._uid,
//                 _created: elementData._created,
//                 _updated: elementData._updated,
//             };
//             if (typeof elementData._deleted === 'boolean') {
//                 returnObj._deleted = elementData._deleted
//             } else {
//                 returnObj._deleted = elementData._deleted === 'TRUE' ? true : false
//             }
//             if (elementData._did !== undefined) returnObj._did = elementData._did;
//             if (elementData._pid !== undefined) returnObj._pid = elementData._pid;
//             if (elementData._eid !== undefined) returnObj._eid = elementData._eid;
//             if (elementData._tid !== undefined) returnObj._tid = elementData._tid;
//             if (elementData._lbl !== undefined) returnObj._lbl = elementData._lbl;
//             if (elementData._emoji !== undefined) returnObj._emoji = elementData._emoji;
//             if (elementData._desc !== undefined) returnObj._desc = elementData._desc;
//             if (elementData._scope !== undefined) returnObj._scope = elementData._scope;
//             if (elementData._type !== undefined) returnObj._type = elementData._type;
//             if (elementData._rollup !== undefined) returnObj._rollup = elementData._rollup;
//             if (elementData._period !== undefined) returnObj._period = parsePeriod(elementData._period);
//             if (elementData._note !== undefined) returnObj._note = elementData._note;
//             if (elementData._source !== undefined) returnObj._source = elementData._source;
//             if (elementData._val !== undefined) returnObj._val = elementData._val;

//             return returnObj;
//         }

//         function parsePeriod(period: any): string {
//             if (typeof period === 'string') return period
//             if (typeof period === 'number') {
//                 period = Math.round(period * 10000) / 10000
//                 try {
//                     return Temporal.Instant.fromEpochMilliseconds((period - (25567 + 1)) * 86400 * 1000).toZonedDateTimeISO(Temporal.Now.timeZone()).toPlainDate().toString();
//                 } catch (e) {
//                     console.log('shit');

//                 }
//             }
//             throw new Error('Unhandled period val...')
//         }
// }

// export function importFirestore(filepath: string): pdw.CompleteDataset {
//     function xlateDate(oldDate: string): pdw.EpochStr {
//         if (typeof oldDate !== 'string') {
//             return pdw.makeEpochStr()
//         }
//         oldDate = oldDate.substring(0, 19) + '+00:00[UTC]'
//         let temp = Temporal.ZonedDateTime.from(oldDate).withTimeZone('America/Chicago');
//         const epoch = pdw.makeEpochStrFromTemporal(temp);
//         return epoch
//     }

//     function xlateScope(oldScope: string): pdw.Scope {
//         if (oldScope === 'Time') return pdw.Scope.SECOND
//         if (oldScope === 'Day') return pdw.Scope.DAY
//         throw new Error('I guess this did happen?')
//     }

//     /**
//      * This function is for firestore, so I declared it in here.
//      * I should do this more.
//      */
//     function parseDef(dataIn: any): pdw.DefLike {
//         let returnDef: pdw.DefLike = {
//             _did: dataIn._did,
//             _lbl: dataIn._lbl,
//             _desc: dataIn._desc,
//             _emoji: dataIn._emoji,
//             _scope: xlateScope(dataIn._scope),
//             _uid: pdw.makeUID(),
//             _deleted: dataIn._deleted,
//             _created: xlateDate(dataIn._created),
//             _updated: xlateDate(dataIn._updated),
//         }

//         return returnDef
//     }

//     function parsePointDef(dataIn: any, defIn: any): pdw.PointDefLike {
//         let pointDef: pdw.PointDefLike = {
//             _did: dataIn._did,
//             _lbl: dataIn._lbl,
//             _desc: dataIn._desc,
//             _emoji: dataIn._emoji,
//             _uid: pdw.makeUID(),
//             _deleted: dataIn._deleted,
//             _created: xlateDate(defIn._created),
//             _updated: xlateDate(defIn._updated),
//             _pid: dataIn._pid,
//             _type: parseType(dataIn._type),
//             _rollup: pdw.Rollup.COUNT
//         }

//         return pointDef

//         function parseType(inData: string): pdw.PointType {
//             if (inData === 'String') return pdw.PointType.TEXT
//             if (inData === 'Boolean') return pdw.PointType.BOOL
//             if (inData === 'Array') return pdw.PointType.SELECT
//             if (inData === 'Enum') return pdw.PointType.SELECT
//             if (inData === 'Number') return pdw.PointType.NUMBER
//             throw new Error('I guess tehre are more')
//         }
//     }

//     function parseEntry(dataIn: any): pdw.EntryLike {
//         //want to update the eid to something for points to reference later
//         dataIn._eid = pdw.makeUID()
//         let entry: pdw.EntryLike = {
//             _eid: dataIn._eid,
//             _note: dataIn._note,
//             _period: parsePeriod(dataIn._period),
//             _did: dataIn._did,
//             _source: dataIn._source,
//             _uid: pdw.makeUID(),
//             _deleted: dataIn._deleted,
//             _created: xlateDate(dataIn._created),
//             _updated: xlateDate(dataIn._updated),
//         }

//         return entry

//         function parsePeriod(text: string): pdw.PeriodStr {
//             if (text.length > 11) {
//                 text = text.substring(0, 19) + '+00:00[UTC]'
//                 let temp = Temporal.ZonedDateTime.from(text).withTimeZone('America/Chicago');
//                 return temp.toPlainDateTime().toString();
//             }
//             if (text.length == 10) return text
//             console.log(text);
//             throw new Error('whatever')

//         }
//     }

//     function parseEntryPoint(point: any, entry: any): pdw.EntryPointLike {
//         let returnPoint: pdw.EntryPointLike = {
//             _eid: entry._eid,
//             _did: entry._did,
//             _pid: point._pid,
//             _val: point._val,
//             _uid: pdw.makeUID(),
//             _deleted: entry._deleted,
//             _created: xlateDate(entry._created),
//             _updated: xlateDate(entry._updated),
//         }
//         return returnPoint
//     }

//     const file = JSON.parse(fs.readFileSync(filepath).toString());
//     const returnData: pdw.CompleteishDataset = {
//         defs: file.defs,
//         pointDefs: file.pointDefs,
//         entries: file.entries,
//         entryPoints: file.entryPoints,
//         tagDefs: file.tagDefs,
//         tags: file.tags
//     }
//     const pdwRef = pdw.PDW.getInstance();

//     let parsedDefs = file.definitions.map((defData: any) => parseDef(defData));
//     let parsedPointDefs: pdw.PointDefLike[] = [];

//     file.definitions.forEach((defData: any) => {
//         let pointDefs: pdw.PointDefLike[] = defData._points.map((pd: any) => parsePointDef(pd, defData))
//         parsedPointDefs.push(...pointDefs)
//     });

//     let parsedEntries = file.entries.map((entryData: any) => parseEntry(entryData));
//     //don't import deleted stuff, it's the only spot they exist
//     parsedEntries = parsedEntries.filter((e: pdw.EntryLike) => e._deleted === false);

//     let parsedEntryPoints: pdw.EntryPointLike[] = [];
//     file.entries.forEach((entry: any) => {
//         if (entry._points === undefined) return
//         let points = entry._points.map((point: any) => parseEntryPoint(point, entry));
//         parsedEntryPoints.push(...points);
//     })

//     parsedEntryPoints = parsedEntryPoints.filter((e: pdw.EntryPointLike) => e._deleted === false);

//     returnData.defs = parsedDefs;
//     returnData.pointDefs = parsedPointDefs;
//     returnData.entries = parsedEntries;
//     returnData.entryPoints = parsedEntryPoints;

//     if (returnData.defs !== undefined) pdwRef.setDefs(parsedDefs);
//     if (returnData.pointDefs !== undefined) pdwRef.setPointDefs(parsedPointDefs);
//     if (returnData.entries !== undefined) pdwRef.setEntries(parsedEntries);
//     if (returnData.entryPoints !== undefined) pdwRef.setEntryPoints(parsedEntryPoints);
//     //just not going to worry about importing tags.
//     if (returnData.tagDefs !== undefined) pdwRef.setTagDefs([]);
//     if (returnData.tags !== undefined) pdwRef.setTags([]);

//     returnData.overview = {
//         storeName: filepath,
//         defs: {
//             current: returnData.defs?.filter(element => element._deleted === false).length,
//             deleted: returnData.defs?.filter(element => element._deleted).length
//         },
//         pointDefs: {
//             current: returnData.pointDefs?.filter(element => element._deleted === false).length,
//             deleted: returnData.pointDefs?.filter(element => element._deleted).length
//         },
//         entries: {
//             current: returnData.entries?.filter(element => element._deleted === false).length,
//             deleted: returnData.entries?.filter(element => element._deleted).length
//         },
//         entryPoints: {
//             current: returnData.entryPoints?.filter(element => element._deleted === false).length,
//             deleted: returnData.entryPoints?.filter(element => element._deleted).length
//         },
//         tagDefs: {
//             current: returnData.tagDefs?.filter(element => element._deleted === false).length,
//             deleted: returnData.tagDefs?.filter(element => element._deleted).length
//         },
//         tags: {
//             current: returnData.tags?.filter(element => element._deleted === false).length,
//             deleted: returnData.tags?.filter(element => element._deleted).length
//         },
//         lastUpdated: pdw.PDW.getDatasetLastUpdate(returnData)
//     }

//     return returnData;
// }

export function importMongo(filepath: string): pdw.CompleteDataset {
    function xlateDate(oldDate: string): pdw.EpochStr {
        if (typeof oldDate !== 'string') {
            return pdw.makeEpochStr()
        }
        oldDate = oldDate.substring(0, 19) + '+00:00[UTC]'
        let temp = Temporal.ZonedDateTime.from(oldDate).withTimeZone('America/Chicago');
        const epoch = pdw.makeEpochStrFrom(temp)!;
        return epoch
    }

    function parseEntry(dataIn: any, defIn: pdw.Def, pointMap: any): pdw.EntryData {
        //want to update the eid to something for points to reference later
        dataIn._eid = pdw.makeUID()
        let entry: pdw.EntryData = {
            _eid: pdw.makeUID(),
            _note: dataIn.note,
            _period: parsePeriod(dataIn.period),
            _did: defIn.did,
            _source: dataIn.source,
            _uid: pdw.makeUID(),
            _deleted: dataIn.deleted,
            _created: xlateDate(dataIn.created),
            _updated: xlateDate(dataIn.updated),
        }
        if(dataIn.hasOwnProperty('points')){
            let oldPids = Object.keys(dataIn.points);
            oldPids.forEach(op=>{
                entry[pointMap[op]] = dataIn.points[op];
            })
        }
        return entry

        function parsePeriod(text: string): pdw.PeriodStr {
            if (text.length > 11) {
                text = text.substring(0, 19) + '+00:00[UTC]'
                let temp = Temporal.ZonedDateTime.from(text).withTimeZone('America/Chicago');
                return temp.toPlainDateTime().toString();
            }
            const parts = text.split('/');
            if(parts.length > 1){
                return parts[2] + '-' + parts[0].padStart(2, '0') + '-' + parts[1].padStart(2, '0');
            }
            if (text.length == 10) return text
            throw new Error('whatever')

        }
    }

    const file = JSON.parse(fs.readFileSync(filepath).toString());
    const returnData: pdw.CompleteDataset = {
        defs: [],
        entries: [],
    }
    const pdwRef = pdw.PDW.getInstance();

    let count = 0;

    file.data.forEach((def: any) => {
        const matchedDef = pdwRef.manifest.find(d=>d.lbl === def.label)
        
        //skip some
        if(def.label === "Reminders Completed") return
        if(def.label === "At Home State") return

        
        if (matchedDef === undefined) throw new Error('No def with label ' + def.label);

        const pointMap = mapPoints(def, matchedDef);
        
        const defEntries = def.entries.map((e:any)=>parseEntry(e,matchedDef, pointMap));

        function mapPoints(def: any, matchedDef: pdw.Def){
            let pointMap:any = {};
            def.points.forEach((dp:any)=>{
                let matchedPoint = matchedDef.data._pts.find(p=>p._lbl === dp.label)!;
                if(matchedPoint !== undefined) return pointMap[dp.pid] = matchedPoint._pid;
                console.log(def.label + " point " + dp.label + ' found no match');
            })
            return pointMap;
        }
        
        // def.points.forEach((pd: any) => {
        //     //parse pointDef
        //     returnData.pointDefs?.push(parsePointDef(pd, def))
        // })

        // if (def.entries === undefined) throw new Error('tat hppened')
        // def.entries.forEach((ent: any) => {
        //     //parse entry
        //     returnData.entries?.push(parseEntry(ent, def))

        //     if (ent.points === undefined) {
        //         count = count + 1
        //         console.warn(ent.mid)
        //     } else {
        //         Object.keys(ent.points).forEach((ep: any) => {
        //             //parse entry points
        //             returnData.entryPoints?.push(parseEntryPoint(ent.points[ep], ent, ep, def))
        //         })
        //     }
        // })
        returnData.entries.push(...defEntries);
    })

    if (returnData.entries !== undefined) pdwRef.setEntries(returnData.entries);
    return returnData;
}

export function importFirestore(filepath: string): pdw.CompleteDataset {
    function xlateDate(oldDate: string): pdw.EpochStr {
        if (typeof oldDate !== 'string') {
            return pdw.makeEpochStr()
        }
        oldDate = oldDate.substring(0, 19) + '+00:00[UTC]'
        let temp = Temporal.ZonedDateTime.from(oldDate).withTimeZone('America/Chicago');
        const epoch = pdw.makeEpochStrFrom(temp)!;
        return epoch
    }

    function parseEntry(dataIn: any, defIn: pdw.Def, pointMap: any): pdw.EntryData {
        //want to update the eid to something for points to reference later
        dataIn._eid = pdw.makeUID()
        let entry: pdw.EntryData = {
            _eid: pdw.makeUID(),
            _note: dataIn.note,
            _period: parsePeriod(dataIn.period),
            _did: defIn.did,
            _source: dataIn.source,
            _uid: pdw.makeUID(),
            _deleted: dataIn.deleted,
            _created: xlateDate(dataIn.created),
            _updated: xlateDate(dataIn.updated),
        }
        if(dataIn.hasOwnProperty('points')){
            let oldPids = Object.keys(dataIn.points);
            oldPids.forEach(op=>{
                entry[pointMap[op]] = dataIn.points[op];
            })
        }
        return entry

        function parsePeriod(text: string): pdw.PeriodStr {
            if (text.length > 11) {
                text = text.substring(0, 19) + '+00:00[UTC]'
                let temp = Temporal.ZonedDateTime.from(text).withTimeZone('America/Chicago');
                return temp.toPlainDateTime().toString();
            }
            const parts = text.split('/');
            if(parts.length > 1){
                return parts[2] + '-' + parts[0].padStart(2, '0') + '-' + parts[1].padStart(2, '0');
            }
            if (text.length == 10) return text
            throw new Error('whatever')

        }
    }

    const file = JSON.parse(fs.readFileSync(filepath).toString());
    const returnData: pdw.CompleteDataset = {
        defs: [],
        entries: [],
    }
    const pdwRef = pdw.PDW.getInstance();

    let count = 0;

    file.definitions.forEach((def: any) => {
        const matchedDef = pdwRef.manifest.find(d=>d.did === def._did)
        
        //skip some
        if(def.label === "Reminders Completed") return
        if(def.label === "At Home State") return

        
        if (matchedDef === undefined){
            throw new Error('No def with label ' + def.label);
        }

        const pointMap = mapPoints(def, matchedDef);
        
        const defEntries = file.entries.filter((ent:any)=>ent._did === def._did)//.map((e:any)=>parseEntry(e,matchedDef, pointMap));
        const processedEntries = defEntries.map((e:any)=>flattenEntry(e, matchedDef.scope));

        console.log(processedEntries);
        
        function mapPoints(def: any, matchedDef: pdw.Def){
            let pointMap:any = {};
            def._points.forEach((dp:any)=>{
                let matchedPoint = matchedDef.data._pts.find(p=>p._pid === dp._pid)!;
                if(matchedPoint !== undefined) return pointMap[dp.pid] = matchedPoint._pid;
                console.log(def.label + " point " + dp.label + ' found no match');
            })
            return pointMap;
        }
        
        function flattenEntry(e: any, scope: pdw.Scope){
            //#TODO - fix created, updated period, kill periodEnd
            let te = def;
            
            e._created = pdw.makeEpochStrFrom(e._created);
            e._updated = pdw.makeEpochStrFrom(e._updated);
            if(e._period.split('.').length == 2) e._period = e._period.split('.')[0] + 'Z';
            e._period = new pdw.Period(pdw.parseTemporalFromEpochStr(pdw.makeEpochStrFrom(e._period!)!).toPlainDateTime().toString(), def.scope).toString();
            // e._period = new Period()
            e._points.forEach((p: any)=>{
                e[p._pid] = p._val
            })
            if(e.hasOwnProperty('_periodEnd')) delete e._periodEnd
            if(e.hasOwnProperty('_points')) delete e._points
            if(e.hasOwnProperty('_next')) delete e._next
            if(e.hasOwnProperty('_prev')) delete e._prev
            return e
        }
        // def.points.forEach((pd: any) => {
        //     //parse pointDef
        //     returnData.pointDefs?.push(parsePointDef(pd, def))
        // })

        // if (def.entries === undefined) throw new Error('tat hppened')
        // def.entries.forEach((ent: any) => {
        //     //parse entry
        //     returnData.entries?.push(parseEntry(ent, def))

        //     if (ent.points === undefined) {
        //         count = count + 1
        //         console.warn(ent.mid)
        //     } else {
        //         Object.keys(ent.points).forEach((ep: any) => {
        //             //parse entry points
        //             returnData.entryPoints?.push(parseEntryPoint(ent.points[ep], ent, ep, def))
        //         })
        //     }
        // })
        returnData.entries.push(...defEntries);
    })

    if (returnData.entries !== undefined) pdwRef.setEntries(returnData.entries);
    return returnData;
}

export function importOldest(filepath: string) {
    console.log('loading...');
    let returnData: pdw.CompleteDataset = {
        defs: [],
        entries: []
    }
    XLSX.set_fs(fs);
    let loadedWb = XLSX.readFile(filepath, { dense: true });
    const shts = loadedWb.SheetNames;
    const pdwRef = pdw.PDW.getInstance();
    const rows: any[][] = ''// loadedWb.Sheets[shts[0]]['!data']!;
    const json = XLSX.utils.sheet_to_json(loadedWb.Sheets[shts[0]]);

    const perRow: { defs: pdw.Def[], pointValMap: { [pid: string]: number, pd: any }[] } = {
        defs: [],
        pointValMap: []
    };

    let periodCol: any, noteCol: any, sourceCol: any;

    rows[0].forEach((cell, i) => {
        // if (cell.c === undefined) return
        const commentText = cell.v//.t
        // const commentBits = commentText.split(":")
        // if (commentBits.length == 1) {
        //     console.error('Comment found:', cell.c.t)
        //     throw new Error('only one line of text in comment')
        // }
        const columnSignifier = commentText//commentBits[1].slice(1);
        if (columnSignifier === '_period') {
            if (periodCol !== undefined) throw new Error('two periods');
            periodCol = i;
            return
        }
        if (columnSignifier === '_note') {
            if (noteCol !== undefined) throw new Error('two notes');
            noteCol = i;
            return
        }
        if (columnSignifier === '_source') {
            if (sourceCol !== undefined) throw new Error('two sources');
            sourceCol = i;
            return
        }
        console.log('Finding PointDef for:', columnSignifier);
        const pd: pdw.PointDef = pdwRef.getPointDefs({ includeDeleted: 'no', pid: [columnSignifier] })[0]
        const def = pdwRef.getDefs({ includeDeleted: 'no', did: [pd._did] })[0]
        console.log(pd);
        if (!perRow.defs.some((e: any) => e._did === def._did)) {
            perRow.defs.push(def);
        }
        //@ts-expect-error
        perRow.pointValMap.push({ loc: i, ['pd']: pd });
    })

    let entriesMade = 0;
    let entryPointsMade = 0;

    rows.forEach((row, i) => {
        if (i === 0) return
        let rowEntries: any = {}
        perRow.pointValMap.forEach(pointValMap => {
            // console.log(pointValMap);
            if (row[pointValMap.loc] !== undefined) {
                if (rowEntries[pointValMap.pd._did] === undefined) {
                    // console.log(i);

                    const newEntry = pdwRef.createNewEntry({
                        _did: pointValMap.pd._did,
                        _period: parsePeriod(row[periodCol].v),
                        _note: noteCol === undefined ? '' : row[noteCol],
                        _source: sourceCol === undefined ? '' : row[sourceCol],
                    })
                    entriesMade = entriesMade + 1
                    const eid = newEntry._eid;
                    // console.log(newEntry);
                    rowEntries[newEntry._did] = newEntry._eid
                }
                pdwRef.createNewEntryPoint({
                    _did: pointValMap.pd._did,
                    _pid: pointValMap.pd._pid,
                    _eid: rowEntries[pointValMap.pd._did],
                    _val: row[pointValMap.loc].v
                })
                entryPointsMade = entryPointsMade + 1;
            }
        })
    })

    console.log('Made ' + entriesMade + ' entries, with ' + entryPointsMade + ' entryPoints');

    function parsePeriod(period: any): string {
        if (typeof period === 'string') return period
        if (typeof period === 'number') {
            return Temporal.Instant.fromEpochMilliseconds((period - (25567 + 1)) * 86400 * 1000).toZonedDateTimeISO(Temporal.Now.timeZone()).toPlainDate().toString();
        }
        throw new Error('Unhandled period val...')
    }
}

export function importOldV9(filepath: string) {
    console.log('loading...');
    let returnData: pdw.CompleteishDataset = {}
    XLSX.set_fs(fs);
    let loadedWb = XLSX.readFile(filepath, { dense: true });
    const shts = loadedWb.SheetNames;
    const pdwRef = pdw.PDW.getInstance();
    const rows: any[][] = loadedWb.Sheets[shts[0]]['!data']!;

    const perRow: { defs: pdw.Def[], pointValMap: { [pid: string]: number, pd: any }[] } = {
        defs: [],
        pointValMap: []
    };

    let periodCol: any, noteCol: any, sourceCol: any;

    let entriesMade = 0;
    let entryPointsMade = 0;

    rows.forEach((row, i) => {
        if (i === 0) return
        const assPointDef = pdwRef.getPointDefs({pid:row[2].v})[0]
        const assDef = pdwRef.getDefs({did: [assPointDef._did]})[0]
        let rowEntries: any = {}
        const newEntry = pdwRef.createNewEntry({
            _did: assDef._did,
            _period: parsePeriod(row[1].v),
            _note: row[4] === undefined ? '' : row[4].v,
            _source: row[0].v,
        })
        entriesMade = entriesMade + 1
        const eid = newEntry._eid;
        pdwRef.createNewEntryPoint({
            _did: assDef._did,
            _pid: assPointDef._pid,
            _eid: newEntry._eid,
            _val: row[3].v
        })
        entryPointsMade = entryPointsMade + 1;

        // console.log('Made ' + entriesMade + ' entries, with ' + entryPointsMade + ' entryPoints');

        function parsePeriod(period: any): string {
            if (typeof period === 'string') return period
            if (typeof period === 'number') {
                period = Math.round(period * 10000)/10000
                return Temporal.Instant.fromEpochMilliseconds(Math.round((period - (25567 + 1)) * 86400 * 1000)).toZonedDateTimeISO(Temporal.Now.timeZone()).toPlainDate().toString();
            }
            throw new Error('Unhandled period val...')
        }
    })
}
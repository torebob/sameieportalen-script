/**
 * Database.gs - Abstraksjonslag for database-operasjoner
 *
 * Dette gjør det enkelt å bytte fra Google Sheets til Firebase senere.
 * Bruk DB.query(), DB.insert(), DB.update(), DB.delete() i stedet for direkte Sheets-kall.
 *
 * Migrering til Firebase: Byt bare "new SheetsProvider()" med "new FirestoreProvider()"
 */

// Global database-instans
const DB = new DatabaseAdapter(new SheetsProvider());

/**
 * DatabaseAdapter - Hovedklasse for alle database-operasjoner
 */
class DatabaseAdapter {
    constructor(provider) {
        this.provider = provider;
    }

    /**
     * Hent data fra database
     * @param {string} collection - Navn på collection (Sheets-navn)
     * @param {object} filters - Filtreringsvilkår {field: value}
     * @returns {Array} Liste med objekter
     */
    query(collection, filters = {}) {
        return this.provider.query(collection, filters);
    }

    /**
     * Legg til nytt dokument
     * @param {string} collection - Navn på collection
     * @param {object} data - Data som skal lagres
     * @returns {object} Lagret data med id
     */
    insert(collection, data) {
        if (!data.id) {
            data.id = Utilities.getUuid();
        }
        return this.provider.insert(collection, data);
    }

    /**
     * Oppdater eksisterende dokument
     * @param {string} collection - Navn på collection
     * @param {string} id - ID på dokumentet
     * @param {object} data - Nye data
     * @returns {boolean} True hvis vellykket
     */
    update(collection, id, data) {
        return this.provider.update(collection, id, data);
    }

    /**
     * Slett dokument
     * @param {string} collection - Navn på collection
     * @param {string} id - ID på dokumentet
     * @returns {boolean} True hvis vellykket
     */
    delete(collection, id) {
        return this.provider.delete(collection, id);
    }

    /**
     * Hent ett enkelt dokument basert på ID
     * @param {string} collection - Navn på collection
     * @param {string} id - ID på dokumentet
     * @returns {object|null} Dokumentet eller null
     */
    getById(collection, id) {
        const results = this.query(collection, { id: id });
        return results.length > 0 ? results[0] : null;
    }
}

/**
 * SheetsProvider - Implementasjon for Google Sheets
 */
class SheetsProvider {
    constructor() {
        this.sheetId = DB_SHEET_ID;
        this.schemas = this._getSchemas();
    }

    /**
     * Definer skjema for alle collections
     * Dette gjør det lett å holde oversikt og endre struktur
     */
    _getSchemas() {
        return {
            'Bookings': ['id', 'resourceId', 'startTime', 'endTime', 'userEmail', 'userName', 'createdAt', 'status'],
            'CommonResources': ['id', 'name', 'description', 'maxBookingHours', 'price', 'cancellationDeadline'],
            'Users': ['email', 'name', 'role', 'apartmentId', 'phone', 'createdAt'],
            'Documents': ['id', 'title', 'url', 'description', 'uploadedBy', 'uploadedAt'],
            'News': ['id', 'title', 'content', 'publishedDate', 'author'],
            'WebsitePages': ['pageId', 'title', 'content', 'password'],
            'AuditLog': ['timestamp', 'userEmail', 'action', 'resource', 'details']
        };
    }

    /**
     * Hent eller opprett sheet
     */
    _getSheet(collection) {
        const ss = SpreadsheetApp.openById(this.sheetId);
        let sheet = ss.getSheetByName(collection);

        if (!sheet) {
            sheet = ss.insertSheet(collection);
            const headers = this.schemas[collection] || ['id', 'data'];
            sheet.appendRow(headers);
        }

        return sheet;
    }

    /**
     * Konverter rad til objekt
     */
    _rowToObject(headers, row) {
        const obj = {};
        headers.forEach((header, i) => {
            obj[header] = row[i];
        });
        return obj;
    }

    /**
     * Konverter objekt til rad
     */
    _objectToRow(headers, obj) {
        return headers.map(header => obj[header] !== undefined ? obj[header] : '');
    }

    /**
     * Query-operasjon
     */
    query(collection, filters = {}) {
        try {
            const sheet = this._getSheet(collection);
            const data = sheet.getDataRange().getValues();

            if (data.length === 0) return [];

            const headers = data.shift();

            let results = data.map(row => this._rowToObject(headers, row));

            // Filtrer resultater
            Object.keys(filters).forEach(key => {
                results = results.filter(item => {
                    // Håndter null/undefined
                    if (filters[key] === null || filters[key] === undefined) {
                        return item[key] === null || item[key] === undefined || item[key] === '';
                    }
                    return item[key] === filters[key];
                });
            });

            return results;
        } catch (e) {
            console.error(`Error in query ${collection}: ${e.message}`);
            throw e;
        }
    }

    /**
     * Insert-operasjon
     */
    insert(collection, data) {
        try {
            const sheet = this._getSheet(collection);
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

            // Legg til timestamp hvis ikke finnes
            if (headers.includes('createdAt') && !data.createdAt) {
                data.createdAt = new Date().toISOString();
            }

            const row = this._objectToRow(headers, data);
            sheet.appendRow(row);

            return data;
        } catch (e) {
            console.error(`Error in insert ${collection}: ${e.message}`);
            throw e;
        }
    }

    /**
     * Update-operasjon
     */
    update(collection, id, data) {
        try {
            const sheet = this._getSheet(collection);
            const sheetData = sheet.getDataRange().getValues();
            const headers = sheetData.shift();
            const idIndex = headers.indexOf('id') !== -1 ? headers.indexOf('id') : headers.indexOf(Object.keys(data)[0]);

            if (idIndex === -1) {
                throw new Error(`Cannot find ID column in ${collection}`);
            }

            const rowIndex = sheetData.findIndex(row => row[idIndex] == id);

            if (rowIndex === -1) {
                throw new Error(`Document with id ${id} not found in ${collection}`);
            }

            // Oppdater kun de feltene som er spesifisert
            headers.forEach((header, colIndex) => {
                if (data[header] !== undefined) {
                    sheet.getRange(rowIndex + 2, colIndex + 1).setValue(data[header]);
                }
            });

            return true;
        } catch (e) {
            console.error(`Error in update ${collection}: ${e.message}`);
            throw e;
        }
    }

    /**
     * Delete-operasjon
     */
    delete(collection, id) {
        try {
            const sheet = this._getSheet(collection);
            const data = sheet.getDataRange().getValues();
            const headers = data[0];
            const idIndex = headers.indexOf('id') !== -1 ? headers.indexOf('id') : 0;

            const rowIndex = data.findIndex(row => row[idIndex] == id);

            if (rowIndex > 0) {
                sheet.deleteRow(rowIndex + 1);
                return true;
            }

            return false;
        } catch (e) {
            console.error(`Error in delete ${collection}: ${e.message}`);
            throw e;
        }
    }
}

/**
 * FirestoreProvider - Placeholder for fremtidig Firebase-migrering
 *
 * Når du er klar for Firebase, implementer disse metodene:
 */
class FirestoreProvider {
    constructor() {
        // Firebase config vil komme her
        throw new Error("FirestoreProvider ikke implementert ennå. Bruk SheetsProvider.");
    }

    async query(collection, filters) {
        // Firebase query-logikk
        // let query = firebase.firestore().collection(collection);
        // Object.keys(filters).forEach(key => {
        //     query = query.where(key, '==', filters[key]);
        // });
        // const snapshot = await query.get();
        // return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    }

    async insert(collection, data) {
        // Firebase insert-logikk
        // const docRef = await firebase.firestore().collection(collection).add(data);
        // return { id: docRef.id, ...data };
    }

    async update(collection, id, data) {
        // Firebase update-logikk
        // await firebase.firestore().collection(collection).doc(id).update(data);
        // return true;
    }

    async delete(collection, id) {
        // Firebase delete-logikk
        // await firebase.firestore().collection(collection).doc(id).delete();
        // return true;
    }
}
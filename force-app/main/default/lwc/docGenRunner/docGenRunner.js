import { LightningElement, api, wire, track } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getTemplatesForObject from '@salesforce/apex/DocGenController.getTemplatesForObject';
import processAndReturnDocument from '@salesforce/apex/DocGenController.processAndReturnDocument';
import generatePdfAsync from '@salesforce/apex/DocGenController.generatePdfAsync';
import checkPdfResult from '@salesforce/apex/DocGenController.checkPdfResult';
import saveGeneratedDocument from '@salesforce/apex/DocGenController.saveGeneratedDocument';

const PDF_POLL_INTERVAL = 2000; // 2 seconds
const PDF_POLL_MAX_ATTEMPTS = 30; // 60 seconds max

export default class DocGenRunner extends LightningElement {
    @api recordId;
    @api objectApiName;

    @track templateOptions = [];
    @track selectedTemplateId;
    @track outputMode = 'download';
    @track templateOutputFormat = 'Document';

    isLoading = false;
    error;
    _templateData = [];
    _pollTimer;

    get outputOptions() {
        const formatLabel = this.templateOutputFormat || 'Document';
        return [
            { label: `Download ${formatLabel}`, value: 'download' },
            { label: `Save to Record (${formatLabel})`, value: 'save' }
        ];
    }

    @wire(getTemplatesForObject, { objectApiName: '$objectApiName' })
    wiredTemplates({ error, data }) {
        if (data) {
            this._templateData = data;
            this.templateOptions = data.map(t => ({
                label: t.Name + (t.Is_Default__c ? ' ★' : ''),
                value: t.Id
            }));
            this.error = undefined;

            // Auto-select default template (first with Is_Default__c = true)
            if (!this.selectedTemplateId) {
                const defaultTemplate = data.find(t => t.Is_Default__c);
                if (defaultTemplate) {
                    this.selectedTemplateId = defaultTemplate.Id;
                    this.templateOutputFormat = defaultTemplate.Output_Format__c || 'Document';
                }
            }
        } else if (error) {
            this.error = 'Error fetching templates: ' + (error.body ? error.body.message : error.message);
            this.templateOptions = [];
        }
    }

    handleTemplateChange(event) {
        this.selectedTemplateId = event.detail.value;
        this.error = null;
        const selected = this._templateData.find(t => t.Id === this.selectedTemplateId);
        if (selected) {
            this.templateOutputFormat = selected.Output_Format__c || 'Document';
        }
    }

    handleOutputModeChange(event) {
        this.outputMode = event.detail.value;
    }

    get isGenerateDisabled() {
        return !this.selectedTemplateId || this.isLoading;
    }

    disconnectedCallback() {
        this._clearPollTimer();
    }

    _clearPollTimer() {
        if (this._pollTimer) {
            // eslint-disable-next-line @lwc/lwc/no-async-operation
            clearInterval(this._pollTimer);
            this._pollTimer = null;
        }
    }

    async generateDocument() {
        this.isLoading = true;
        this.error = null;

        try {
            const selected = this._templateData.find(t => t.Id === this.selectedTemplateId);
            const templateType = selected ? selected.Type__c : 'Word';
            const isPPT = templateType === 'PowerPoint';
            const isExcel = templateType === 'Excel';
            const isPDF = this.templateOutputFormat === 'PDF' && !isPPT && !isExcel;

            if (isPDF) {
                this.showToast('Info', 'Generating PDF...', 'info');
                const saveToRecord = this.outputMode === 'save';

                const result = await generatePdfAsync({
                    templateId: this.selectedTemplateId,
                    recordId: this.recordId,
                    saveToRecord: saveToRecord
                });

                if (result.saved) {
                    this.showToast('Success', 'PDF saved to record.', 'success');
                } else if (result.base64) {
                    const docTitle = result.title || 'Document';
                    this.downloadBase64(result.base64, docTitle + '.pdf', 'application/pdf');
                    this.showToast('Success', 'PDF downloaded.', 'success');
                }
                this.isLoading = false;
            } else {
                // Native DOCX/PPTX path
                const result = await processAndReturnDocument({
                    templateId: this.selectedTemplateId,
                    recordId: this.recordId
                });

                if (!result || !result.base64) {
                    throw new Error('Document generation returned empty result.');
                }

                const ext = isExcel ? 'xlsx' : (isPPT ? 'pptx' : 'docx');
                const docTitle = result.title || 'Document';

                if (this.outputMode === 'save') {
                    this.showToast('Info', 'Saving to Record...', 'info');
                    await saveGeneratedDocument({
                        recordId: this.recordId,
                        fileName: docTitle,
                        base64Data: result.base64,
                        extension: ext
                    });
                    this.showToast('Success', `${ext.toUpperCase()} saved to record.`, 'success');
                } else {
                    this.downloadBase64(result.base64, docTitle + '.' + ext, 'application/octet-stream');
                    const typeLabel = isExcel ? 'Excel spreadsheet' : (isPPT ? 'PowerPoint' : 'Word document');
                    this.showToast('Success', `${typeLabel} downloaded.`, 'success');
                }
                this.isLoading = false;
            }
        } catch (e) {
            let msg = 'Unknown error during generation';
            if (e.body && e.body.message) {
                msg = e.body.message;
            } else if (e.message) {
                msg = e.message;
            } else if (typeof e === 'string') {
                msg = e;
            }
            this.error = 'Generation Error: ' + msg;
            this.isLoading = false;
        }
    }

    /**
     * Polls for async PDF generation result and triggers download when ready.
     */
    _pollForPdfResult(resultKey, docTitle) {
        let attempts = 0;
        // eslint-disable-next-line @lwc/lwc/no-async-operation
        this._pollTimer = setInterval(async () => {
            attempts++;
            try {
                const result = await checkPdfResult({ resultKey: resultKey });
                if (result && result.base64) {
                    this._clearPollTimer();
                    this.downloadBase64(result.base64, docTitle + '.pdf', 'application/pdf');
                    this.showToast('Success', 'PDF downloaded.', 'success');
                    this.isLoading = false;
                } else if (attempts >= PDF_POLL_MAX_ATTEMPTS) {
                    this._clearPollTimer();
                    this.error = 'PDF generation timed out. Please check the record for the generated file.';
                    this.isLoading = false;
                }
            } catch (err) {
                this._clearPollTimer();
                this.error = 'Error checking PDF result: ' + (err.body ? err.body.message : err.message);
                this.isLoading = false;
            }
        }, PDF_POLL_INTERVAL);
    }

    /**
     * Downloads a base64-encoded file via an anchor element.
     */
    downloadBase64(base64Data, fileName, mimeType) {
        const binaryString = atob(base64Data);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        const blob = new Blob([bytes], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    showToast(title, message, variant) {
        this.dispatchEvent(new ShowToastEvent({ title, message, variant }));
    }
}

/**
 * Simple Wizard Plugin System
 * For developers to easily integrate forms with Excel template system
 */

class WizardPlugin {
    constructor(config) {
        this.templateId = config.templateId;
        this.onSubmit = config.onSubmit || this.defaultSubmit;
        this.autofill = config.autofill || {};
        this.dataMapper = config.dataMapper || null;
    }

    // Simple autofill function
    fillForm(data) {
        Object.keys(data).forEach(key => {
            const element = document.getElementById(key) || document.querySelector(`[name="${key}"]`);
            if (element) {
                element.value = data[key];
                element.dispatchEvent(new Event('change'));
            }
        });
    }

    // Submit form using the provided dataMapper function
    submitForm() {
        let formData;
        
        // If dataMapper is provided, use it directly
        if (this.dataMapper && typeof this.dataMapper === 'function') {
            formData = this.dataMapper();
        } else {
            // Fallback: collect basic form data
            const form = document.querySelector('form') || document.body;
            const inputs = form.querySelectorAll('input, select, textarea');
            const rawData = {};

            inputs.forEach(input => {
                const key = input.id || input.name;
                if (key) {
                    if (input.type === 'checkbox') {
                        rawData[key] = input.checked;
                    } else if (input.type === 'radio') {
                        if (input.checked) rawData[key] = input.value;
                    } else if (input.type === 'number') {
                        rawData[key] = parseFloat(input.value) || 0;
                    } else {
                        rawData[key] = input.value;
                    }
                }
            });
            formData = rawData;
        }
        this.onSubmit(formData);
    }

    // Default submit handler
    defaultSubmit(data) {
        // Send to parent window (if in iframe)
        if (window.parent !== window) {
            window.parent.postMessage({
                type: 'wizardFormSubmitted',
                templateId: this.templateId,
                data: data // Send the complete data structure as collected by dataMapper
            }, window.location.origin);
        } else {
            console.log('Form Data:', JSON.stringify(data, null, 2));
        }
    }

    setup() {
        const submitButtons = document.querySelectorAll('button[type="submit"], .submit-btn, #submitBtn');
        submitButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                this.submitForm();
            });
        });

        if (Object.keys(this.autofill).length > 0) {
            this.fillForm(this.autofill);
        }

        this.addTempDataButton();
    }

    addTempDataButton() {
        const tempBtn = document.createElement('button');
        tempBtn.textContent = 'Fill Test Data';
        tempBtn.className = 'temp-data-btn';
        tempBtn.style.cssText = `
            background: #6c757d; color: white; border: none; 
            padding: 8px 16px; border-radius: 4px; margin: 10px;
            cursor: pointer; font-size: 12px;
        `;
        
        tempBtn.addEventListener('click', (e) => {
            e.preventDefault();
            this.fillTestData();
        });

        const target = document.querySelector('form') || document.body;
        target.insertBefore(tempBtn, target.firstChild);
    }

    fillTestData() {
        console.log('Override fillTestData() method for template-specific test data');
    }
}

window.WizardPlugin = WizardPlugin;


async function downloadPDF(fileName) {

    $('#pdfLoader').show();
    try {
        // Store current accordion state
        let accordionStates = [];

        $('.accordion-collapse').each(function () {
            accordionStates.push($(this).hasClass('show'));
        });

        // Open all accordions
        $('.accordion-collapse').addClass('show');

        // Hide accordion buttons/icons if needed
        $('.accordion-button').hide();

        // Hide microphone buttons
        $('.mic-btn').hide();
        $('#btnSection').hide();
        $('#btnSaveSection').hide();

        // Optional: remove borders/shadows for clean PDF
        $('.accordion-item').css({
            'border': 'none',
            'box-shadow': 'none'
        });

        //         $('#patientForm').css({
        //     'font-size': '11px',
        //     'line-height': '1.3'
        // });

        // $('#patientForm input, #patientForm textarea, #patientForm select, #patientForm p, #patientForm label').css({
        //     'font-size': '11px',
        //     'padding': '2px 4px',
        //     'height': 'auto'
        // });

        // $('h4, h5, h6').css({
        //     'font-size': '14px'
        // });

        // Wait for rendering

        $('input[type="date"]').each(function () {
            const originalValue = $(this).val(); // This is always yyyy-mm-dd

            if (originalValue && originalValue.includes('-')) {
                const parts = originalValue.split('-');
                // Convert yyyy-mm-dd to mm-dd-yyyy
                const formattedDate = `${parts[1]}-${parts[2]}-${parts[0]}`;

                // Store original state
                $(this).data('original-date', originalValue);

                // Switch type to text and force the value into the attribute
                $(this).attr('type', 'text');
                $(this).val(formattedDate);
                $(this).attr('value', formattedDate); // Crucial for html2pdf rendering
            }
        });

        // Handle Textareas (Auto-expand height so no scrollbars)
        $('textarea').each(function () {
            const $this = $(this);
            // Save original styles
            $this.data('original-style', $this.attr('style') || '');

            // Create a temporary div to replace it for perfect rendering
            const val = $this.val();
            const $replacement = $('<div class="pdf-temp-text"></div>').text(val).css({
                'width': $this.width() + 'px',
                'min-height': $this.height() + 'px',
                'border': '1px solid #dee2e6',
                'padding': '5px',
                'white-space': 'pre-wrap', // Preserves line breaks
                'word-wrap': 'break-word',
                'background-color': '#fff'
            });
            $this.hide().after($replacement);
        });

        // Handle Text Inputs (Prevent clipping of long text)
        $('input[type="text"]').each(function () {
            const $this = $(this);
            // Ignore the ones we just converted from date (unless they are also very long)
            if ($this.val().length > 15) {
                $this.data('was-text-input', true);
                const $replacement = $('<span class="pdf-temp-text"></span>').text($this.val()).css({
                    'display': 'inline-block',
                    'border-bottom': '1px solid #dee2e6',
                    'min-width': $this.width() + 'px',
                    'padding': '0 5px'
                });
                $this.hide().after($replacement);
            }
        });

        await new Promise(resolve => setTimeout(resolve, 1000));


        // Target form/div
        const element = document.getElementById('patientForm');

        // PDF options
        const opt = {
            margin: 0.3,
            filename: fileName,
            image: {
                type: 'jpeg',
                quality: 1
            },
            html2canvas: {
                scale: 2,
                useCORS: true,
                scrollY: 0
            },
            jsPDF: {
                unit: 'in',
                format: 'a4',
                orientation: 'portrait'
            },
            pagebreak: {
                mode: ['avoid-all', 'css', 'legacy']
            }
        };

        // Generate PDF
        await html2pdf().set(opt).from(element).save();

        $('input[data-original-date]').each(function () {
            const originalDate = $(this).data('original-date');

            $(this).attr('type', 'date');
            $(this).val(originalDate);
            $(this).attr('value', originalDate); // Restore attribute

            $(this).removeData('original-date');
        });

        // Restore accordion state
        $('.accordion-collapse').each(function (index) {

            if (!accordionStates[index]) {
                $(this).removeClass('show');
            }
        });

        // Remove temp text replacements and show original textareas/inputs
        $('.pdf-temp-text').remove();
        $('textarea, input').show();

        //         $('#patientForm').css({
        //     'font-size': '',
        //     'line-height': ''
        // });

        // $('#patientForm input, #patientForm textarea, #patientForm select,#patientForm p, #patientForm label').css({
        //     'font-size': '',
        //     'padding': '',
        //     'height': ''
        // });

        // $('h4, h5, h6').css({
        //     'font-size': ''
        // });

        // Show hidden buttons again
        $('.accordion-button').show();
        $('.mic-btn').show();
        $('#btnSection').show();
        $('#btnSaveSection').show();
    }
    finally {

        // HIDE LOADER
        $('#pdfLoader').hide();
    }
}
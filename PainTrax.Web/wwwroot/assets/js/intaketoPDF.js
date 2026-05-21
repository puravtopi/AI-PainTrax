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
        await new Promise(resolve => setTimeout(resolve, 500));

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

        // Restore accordion state
        $('.accordion-collapse').each(function (index) {

            if (!accordionStates[index]) {
                $(this).removeClass('show');
            }
        });

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
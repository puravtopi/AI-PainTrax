ClassicEditor
    .create(document.querySelector('#txtarea'), {


        mention: {
            feeds: [
                {
                    marker: '@',
                    feed: ['@Barney', '@Lily', '@Marry Ann', '@Marshall', '@Robin', '@Ted'],
                    minimumCharacters: 1
                }
            ]
        }
    })
    .catch(error => {
        console.error(error);
    });
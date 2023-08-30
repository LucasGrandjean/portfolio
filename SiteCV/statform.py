from django import forms

class FileUploadForm(forms.Form):
    csv_file = forms.FileField(label='Fichier CSV', widget=forms.ClearableFileInput(attrs={'accept': '.csv'}))
    choice = forms.ChoiceField(label='Choix', choices=[('bilan', 'Bilan'), ('mediateur', 'Médiateur')])
    
class SiteSelectionForm(forms.Form):
    def __init__(self, *args, site_choices=[], **kwargs):
        super().__init__(*args, **kwargs)

        self.fields['selected_sites'] = forms.MultipleChoiceField(
            choices=[(site, site) for site in site_choices],
            widget=forms.CheckboxSelectMultiple
        )
        
        self.fields['csv_file'] = forms.FileField(
            label='Fichier CSV',
            widget=forms.ClearableFileInput(attrs={'accept': '.csv'}),
            required=True
        )
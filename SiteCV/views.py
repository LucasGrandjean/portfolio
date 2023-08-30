from django.shortcuts import render
from django.http import HttpResponse,FileResponse
import csv
from .statform import FileUploadForm, SiteSelectionForm   #AutoStat
from .statutils import genererNomFichier, formatLieu, doExtract, summary_mediateur, creerBilanSynthese, filter_csv_by_sites #AutoStat
from django import forms
import tempfile
import os
import zipfile
from django.shortcuts import redirect
import shutil
from django.urls import reverse

def index(request):
    return render(request, 'index.html')
    
def education(request):
    return render(request, 'education.html')
    
def contact(request):
    return render(request, 'contact.html')
    
def formation(request):
    return render(request, 'project.html')

class FileUploadForm(forms.Form):
    csv_file = forms.FileField(label='Fichier CSV', widget=forms.ClearableFileInput(attrs={'accept': '.csv'}))
    choice = forms.ChoiceField(label='Choix', choices=[('1', 'Bilan'), ('2', 'Synthèse de Bilan'), ('3', 'Bilan Médiateur')])
 
def site_selection(request, site_choices):
    # Convert the comma-separated string to a list of choices
    site_choices_list = site_choices.split(',') if site_choices else []
    
    if request.method == 'POST':
        form_site_selection = SiteSelectionForm(request.POST, request.FILES, site_choices=site_choices_list)  # Include request.FILES here
        if form_site_selection.is_valid():
            selected_sites = form_site_selection.cleaned_data['selected_sites']
            filtered_data = filter_csv_by_sites(request.FILES['csv_file'], selected_sites)
            output_filename = doExtract(filtered_data, 'Bilan_Synthese', 1)
            with open(output_filename, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                response['Content-Disposition'] = 'attachment; filename="filtered_synthesis.xlsx"'
                return response
    else:
        form_site_selection = SiteSelectionForm(site_choices=site_choices_list)

    return render(request, 'site_selection.html', {'form_site_selection': form_site_selection})
 
def autostat(request):
    if request.method == 'POST':
        form = FileUploadForm(request.POST, request.FILES)
        if form.is_valid() or 'selected_sites' in request.POST:
            csv_file = form.cleaned_data['csv_file']
            choice = form.cleaned_data['choice']

            # Lire les données CSV depuis le fichier téléchargé
            csv_data = csv_file.read().decode('utf-8')
            csv_lines = csv_data.split('\n')
            csv_reader = csv.DictReader(csv_lines, delimiter=';')

            donneesLieux = {}
            for row in csv_reader:
                site = row['Site']
                if site not in donneesLieux:
                    donneesLieux[site] = []
                donneesLieux[site].append(row)

            if choice == '2':
                site_choices = ','.join(donneesLieux.keys())  # Convert list to comma-separated string
                redirect_url = reverse('site_selection', args=[site_choices])
                return redirect(redirect_url)
                        

            elif choice == '3':  # Bilan Médiateur
                merged_data = creerBilanSynthese(donneesLieux)
                output_filename = 'Bilan_Mediateur.xlsx'
                file_mediateur = summary_mediateur(merged_data,output_filename)
                with open(file_mediateur, 'rb') as f:
                    response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
                    return response
            else:  # Inconnu : Synthèse
            
                temp_dir = tempfile.mkdtemp()  # Create a temporary directory

                generated_files = []  # To store paths of generated files
                for site, data in donneesLieux.items():
                    output_filename = genererNomFichier("Bilan", site)
                    output_excel_path = doExtract(data, output_filename, 1)  # Utiliser la variable 'data' à la place de 'merged_data'
                    generated_files.append(output_excel_path)  # Store the path of the generated file

                # Create a zip file containing all generated Excel files
                zip_path = os.path.join(temp_dir, 'generated_files.zip')
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for generated_file in generated_files:
                        zipf.write(generated_file, os.path.basename(generated_file))
                        os.remove(generated_file) 

                # Read the zip file and create a response to return for download
                with open(zip_path, 'rb') as f:
                    response = HttpResponse(f.read(), content_type='application/zip')
                    response['Content-Disposition'] = 'attachment; filename="generated_files.zip"'

                # Clean up the temporary directory
                shutil.rmtree(temp_dir)            
                # Return the responses
                return response
    else:
        form = FileUploadForm()

    return render(request, 'autostat.html', {'form': form})
  
def pix(request):
    file = open('./templates/Certifications/Pix.pdf', 'rb')
    return FileResponse(file)
    
def ccp1(request):
    file = open('./templates/Certifications/Ccp1.pdf', 'rb')
    return FileResponse(file)
    
def citoyen(request):
   file = open('./templates/Certifications/Citoyen.pdf', 'rb')
   return FileResponse(file)
    
def psc1(request):
    file = open('./templates/Certifications/Psc1.pdf', 'rb')
    return FileResponse(file)
    
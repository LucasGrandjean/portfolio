from . import views
from django.urls import path
urlpatterns = [
    path("", views.index, name="index"),
    path("Education/", views.education, name="education"),  
    path("Contact/", views.contact, name="contact"),       
    path("Project/", views.formation, name="formation"), 
    path("Project/AutoStat", views.autostat, name="autostat"),
    path('site_selection/<str:site_choices>/', views.site_selection, name='site_selection'),
    path("Education/PIX", views.pix, name="pix"),  
    path("Education/PSC1", views.psc1, name="psc1"),      
    path("Education/CNFS", views.ccp1, name="ccp1"),  
    path("Education/Citoyen", views.citoyen, name="Citoyen"),
]
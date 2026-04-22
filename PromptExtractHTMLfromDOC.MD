
*I'm building a website about 1920s world music. I have a single-page HTML file (`roaring20s.html`) and a DOCX file (`TheRoaring20s_-_New_Material.docx`) containing vignettes and embedded images. I want to:*

*1. Parse the DOCX XML to extract vignettes (text blocks) and map which images belong to each vignette by document position*
*2. Create `vignettes/<country>_<n>.html` files for each vignette, with images referenced as `images/<country>_<n>_pic<m>.jpg`*
*3. Extract and save the images to the images folder with those names*
*4. Create blank vignette shells for the 22 countries that have no DOCX content yet*
*5. Modify `roaring20s.html` to fetch and inject the vignette HTMLs when a country is selected*

*The countries in the site are: us, cuba, caribbean, argentina, brazil, mexico, colombia, france, germany, ballet, britain, ireland, spain, italy, portugal, russia, hungary, czechia, poland, balkans, japan, china, india, indonesia, philippines, korea, egypt, africa, middleeast, australia*

*Start by doing a recon pass — parse the DOCX XML and give me a manifest of what vignettes exist and which images map to each one, before writing any files.*

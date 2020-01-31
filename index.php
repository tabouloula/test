<?php
    $excel = '';
	if(isset($_POST["submit"])){
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="FICHE_EXP.xlsx"');
		header('Cache-Control: max-age=0');
        require_once 'PHPExcel/Classes/PHPExcel/IOFactory.php';
        require_once 'PHPExcel/Classes/PHPExcel.php';	
        $excel2 = PHPExcel_IOFactory::createReader('Excel2007');
        $excel2 = $excel2->load('ficheExpéditions.xlsx'); // Empty Sheet
        $excel2->setActiveSheetIndex(0);
		for($i=0;$i<8;$i++){
			$cel = 25 + $i;
			$Autre = ($_POST["libelle"][$i] == "Autres".$i ? $_POST["libelle-autre"][$i] : $_POST["libelle"][$i]);
			$excel2->getActiveSheet()
			->setCellValue('A'.$cel,ucfirst($Autre))
			->setCellValue('D'.$cel,ucfirst($_POST["quantiy"][$i]))
			->setCellValue('H'.$cel,ucfirst($_POST["poids"][$i]))
			->setCellValue('I'.$cel,ucfirst($_POST["unite"][$i]))
			->setCellValue('J'.$cel,ucfirst($_POST["total"][$i]));
		}
        $excel2->getActiveSheet()
			->setCellValue('A7', $_POST["adresse-expediteur"])
			->setCellValue('B7', $_POST["nom"]." ".$_POST["prenom"]." ".$_POST["email"])
            ->setCellValue('G7', $_POST["cdf"])
            ->setCellValue('J7', $_POST["date-envoie"])       
            ->setCellValue('B10', $_POST["dangereux"])       
            ->setCellValue('B15', $_POST["urgence"])       
            ->setCellValue('H10', $_POST["classe"])       
            ->setCellValue('B14', $_POST["numero_colis"])       
            ->setCellValue('J16', $_POST["colisage"])       
            ->setCellValue('B13', $_POST["type-envoie"])
            ->setCellValue('B17', $_POST["name"])
            ->setCellValue('H17', $_POST["firstname"])
            ->setCellValue('C18', $_POST["company"])
            ->setCellValue('C19', $_POST["adresse"])
            ->setCellValue('C20', $_POST["zip"])
            ->setCellValue('G20', $_POST["city"])
            ->setCellValue('B21', $_POST["country"])
            ->setCellValue('B22', $_POST["phone"])
            ->setCellValue('I35', $_POST["date-livraison"])
            ->setCellValue('H37', $_POST["commentaire"]);
        $objWriter = PHPExcel_IOFactory::createWriter($excel2, 'Excel2007');
        $objWriter->save('php://output');
    }
?>
<!doctype html>
<html lang="en">
  <head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
	<style>
		.form-control {
			font-size: 14px;
		}
		label {
			font-size: 14px;
		}
		fieldset 
		{
			border: 1px solid #ddd !important;
			margin: 20px 0px 0px 0px;
			xmin-width: 0;
			padding: 10px;       
			position: relative;
			border-radius:4px;
			background-color:#f5f5f5;
			padding-left:10px!important;
		}	
		
		legend
		{
			font-size:14px;
			font-weight:bold;
			margin-bottom: 0px; 
			width: 35%; 
			border: 1px solid #ddd;
			border-radius: 4px; 
			padding: 5px 5px 5px 10px; 
			background-color: #ffffff;
		}
	</style>
    <title>Fiche d'expédition</title>
  </head>
  <body>
      <div class="container">
            <div class="jumbotron">
				<h2 class="text-center">Fiche d'expédition</h2>				
				<small id="emailHelp" class="form-text text-muted text-center">( tous les champs à renseigner sont obligatoires )</small>
				<br />
				<?php echo $excel; ?>
                <form action="index.php" method="post">
                    <fieldset class="col-md-12">    	
						<legend>Expéditeur</legend>
						<div class="form-group">
							<div class="row">
								<div class="col-md-4">
									<label for="exampleInputPassword1">Nom</label>
									<input type="text" name="nom" class="form-control" placeholder="Nom" required>
								</div>
								<div class="col-md-4">
									<label for="exampleInputPassword1">Prénom</label>
									<input type="text" name="prenom" class="form-control" placeholder="Prénom" required>
								</div>
								<div class="col-md-4">
									<label for="exampleInputPassword1">E-mail</label>
									<input type="text" name="email" class="form-control" placeholder="E-mail" required>
								</div>						
							</div>                        
						</div>
						<div class="form-group">
							<div class="row">
								<div class="col-md-12">
									<label for="exampleInputPassword1">Adresse</label>
									<textarea col="12" row="3" name="adresse-expediteur" class="form-control" required>SYMRISE 15 rue Mozart 92 110 Clichy</textarea>
								</div>												
							</div>                        
						</div>
						<div class="form-group">
							<div class="row">
								<div class="col-6">
									<label for="exampleInputPassword1">N° Centre de frais</label>
									<select name="cdf" class="form-control" required>
										<option value="">Cost Center Name</option>
										<option value="FR00625935">BOTANICALS-LABS</option>
										<option value="FR00673060">CI - GLOBAL DIRECTION</option>
										<option value="FR00673050">CI - L'OREAL SALES</option>
									</select>
								</div>
								<div class="col-6">
									<label for="exampleInputPassword1">Date de l'envoie</label>
									<input type="date" name="date-envoie"  class="form-control" placeholder="Date de l'envoie" required>
								</div>
													
							</div>                        
						</div>
					</fieldset>
					<fieldset class="col-md-12">    	
						<legend>Type d'envoie</legend>
						<div class="form-group">
							<div class="row">
								<div class="col-3">
									<label for="exampleInputPassword1">Numéro de colis</label>
									<input type="texte" name="numero_colis" class="form-control" placeholder="Numéro de colis" required>								
								</div>		
								<div class="col-3">
									<label for="exampleInputPassword1">Type d'envoie</label>
									<select name="type-envoie" class="form-control" required>
										<option value="">Selectionner type d'envoie</option>
										<option value="Par coursier">Par coursier</option>
										<option value="Par navette Holzminden">Par navette Holzminden</option>
										<option value="Autre: DHL, ASAP, THE COURIER COMPANY, CHRONOPOST, FEDEX">Autre: DHL, ASAP, THE COURIER COMPANY, CHRONOPOST, FEDEX</option>
									</select>
								</div>																	
								<div class="col-3">
									<label for="exampleInputPassword1">Colisage (L x l x H)</label>
									<select name="colisage" class="form-control" required>
										<option value="">Selectionner colisage (L x l x H)</option>
										<option value="20 x 14 x 14">20 x 14 x 14</option>
										<option value="25 x 20 x 15">25 x 20 x 15</option>
										<option value="30 x 20 x 20">30 x 20 x 20</option>
										<option value="40 x 30 x 20">40 x 30 x 20</option>
										<option value="50 x 40 x 30">50 x 40 x 30</option>
										
									</select>
								</div>
								<div class="col-3">
									<label for="exampleInputPassword1">Degrés d'urgence</label>
									<select name="urgence" class="form-control" required>
										<option value="">Selectionner degrés d'urgence</option>
										<option value="Express 9h">Express 9h</option>
										<option value="Express 10h">Express 10h</option>
										<option value="Express 12h">Express 12h</option>
										
									</select>
								</div>								
							</div>                        
						</div>
						<div class="form-group">
							<div class="row">
								<div class="col-3">
									<div class="custom-control custom-checkbox">
										<input type="radio" name="dangereux" value="Dangereux" class="custom-control-input" id="defaultUnchecked" checked>
										<label class="custom-control-label" for="defaultUnchecked">Dangereux</label>
									</div>								
								</div>
								<div class="col-3">
									<div class="custom-control custom-checkbox">
										<input type="radio" name="dangereux" value="Non dangereux" class="custom-control-input" id="defaultdangereux">
										<label class="custom-control-label" for="defaultdangereux">Non dangereux</label>
									</div>
								</div>
								<div class="col-3">
									<div class="custom-control custom-checkbox">
										<input type="radio" name="classe" value="UN" class="custom-control-input" id="defaultUn" checked>
										<label class="custom-control-label" for="defaultUn">UN</label>
									</div>								
								</div>
								<div class="col-3">
									<div class="custom-control custom-checkbox">
										<input type="radio" name="classe" value="Classe" class="custom-control-input" id="defaultclasse">
										<label class="custom-control-label" for="defaultclasse">Classe</label>
									</div>
								</div>									
							</div>
						</div>						
					</fieldset>
					<fieldset class="col-md-12">    	
						<legend>Adresse de destination / Destination address :</legend>
						<div class="form-group">
							<div class="row">
								<div class="col-6">
									<label for="exampleInputPassword1">Nom / Name </label>
									<input type="text" name="name" class="form-control" placeholder="Nom / Name" required>
								</div>
								<div class="col-6">
									<label for="exampleInputPassword1">Prénom / First name  </label>
									<input type="text" name="firstname" class="form-control" placeholder="Prénom / First name" required>
								</div>													
							</div>
						</div>
						<div class="form-group">
							<div class="row">
								<div class="col-6">
									<label for="exampleInputPassword1">Société / Company </label>
									<input type="text" name="company" class="form-control" placeholder="Société / Company" required>
								</div>
								<div class="col-6">
									<label for="exampleInputPassword1">Adresse / Address </label>
									<input type="text" name="adresse" class="form-control" placeholder="Adresse / Address" required>
								</div>									
							</div> 							
						</div>						
						<div class="form-group">
							<div class="row">
								<div class="col-6">
									<label for="exampleInputPassword1">Code postal  / ZIP code </label>
									<input type="text" name="zip" class="form-control" placeholder="Code postal  / ZIP code" required>
								</div>
								<div class="col-6">
									<label for="exampleInputPassword1">Ville / City  </label>
									<input type="test" name="city" class="form-control" placeholder="Ville / City" required>
								</div>													
							</div>
						</div>
						<div class="form-group">
							<div class="row">
								<div class="col-6">
									<label for="exampleInputPassword1">Pays / Country </label>
									<input type="text" name="country" class="form-control" placeholder="Pays / Country" required>
								</div>
								<div class="col-6">
									<label for="exampleInputPassword1">Telephone  N° </label>
									<input type="text" name="phone" class="form-control" placeholder="Telephone  N°" required>
								</div>									
							</div> 							
						</div>
						
					</fieldset>
					<fieldset class="col-md-12">    	
						<legend>Description du colis / Description of the parcel</legend>
						<div class="form-group">
							<div class="row">
								<div class="col-3">
									<label for="exampleInputPassword1" style="height: 45px;">LIBELLE PRODUIT</label>
									<?php for($i=0; $i<8; $i++) { echo '
									<select name="libelle[]" id="libelle'.$i.'" class="form-control" style="margin-bottom: 10px;">
										<option value="">LIBELLE</option>
										<option value="Affiches">Affiches</option>
										<option value="Arôme Alimentaire">Arôme Alimentaire</option>
										<option value="Autres'.$i.'">Autres</option>
										<option value="Bougies">Bougies</option>
										<option value="Boissons alcoolisées">Boissons alcoolisées</option>
										<option value="Cartes">Cartes</option>
										<option value="Concentré">Concentré</option>
										<option value="Cosmétiques">Cosmétiques</option>
										<option value="Déodorants inflammables">Déodorants inflammables</option>
										<option value="Documents">Documents</option>
										<option value="Mouillettes">Mouillettes</option>
										<option value="Parfum / solution alcoolique">Parfum / solution alcoolique</option>
										<option value="Produits du marché">Produits du marché</option>
									</select> 
									<input type="text" name="libelle-autre[]" class="form-control output'.$i.'" style="display:none;margin-bottom: 10px;" placeholder="Saisir le libelle de produit">';
									
									} ?>
								</div>
								<div class="col-2">
									<label for="exampleInputPassword1" style="height: 45px;">QUANTITE D'ECHANTILLONS </label>
									<?php for($i=0; $i<8; $i++) { 
										echo '<input type="text" name="quantiy[]" class="form-control" style="margin-bottom: 10px;line-height: 25px;">';
									 } ?>
								</div>	
								<div class="col-2">
									<label for="exampleInputPassword1" style="height: 45px;">POIDS / CONTENANCE UNITAIRE</label>
									<?php for($i=0; $i<8; $i++) { 
										echo '<input type="text" name="poids[]" class="form-control" style="margin-bottom: 10px;line-height: 25px;">';
									 } ?>
									
								</div>
								<div class="col-2">
									<label for="exampleInputPassword1" style="height: 45px;">UNITE DE MESURE</label>
									<?php for($i=0; $i<8; $i++) { 
										echo '<input type="text" name="unite[]" class="form-control" style="margin-bottom: 10px;line-height: 25px;">';
									 } ?>
								</div>
								<div class="col-3">
									<label for="exampleInputPassword1" style="height: 45px;">CONTENANCE TOTAL </label>
									<?php for($i=0; $i<8; $i++) { 
										echo '<input type="text" name="total[]" class="form-control" style="margin-bottom: 10px;line-height: 25px;">';
									 } ?>
								</div>									
							</div>                        
						</div>
					</fieldset>
					<fieldset>
						<legend>Commentaires supplémentaires</legend>						
						<div class="form-group">
							<div class="row">
								<div class="col-6">
									<label for="exampleInputPassword1">Délai de livraison</label>
									<input type="date" name="date-livraison" pattern="\d{1,2}/\d{1,2}/\d{4}" class="form-control" placeholder="Délai de livraison" required>
								</div>
								<div class="col-6">
									<label for="exampleInputPassword1">Commentaire supplémentaire</label>
									<textarea col="12" row="3" name="commentaire" class="form-control"></textarea>
								</div>
													
							</div>                        
						</div>				
					</fieldset>
					
						<button type="submit" name="submit" class="btn btn-primary float-right">Envoyer</button>
					
                </form>
				
            </div>
        </div>
		<script type="text/javascript">

			
            $(document).ready(function() {
                $("select").on('change', function() { 
                    $(this).find("option:selected").each(function() { 
                        var geeks = $(this).attr("value");
						for (var i = 0; i<8; i++) {
							if (geeks == "Autres"+i) { 
								$("#libelle"+i).hide(); 
								$(".output"+i).show(); 
							} else { 
								$(".output"+i).hide();
								$("#libelle"+i).show();
							} 
						}
  
                    }); 
                }).change(
				); 
            }); 
        </script>

</body>
</html>
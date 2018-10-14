<!DOCTYPE html>
<html>
<head>
    <title>Print Invoice</title>
    <style>
        .page
        {           
            padding:2cm;
        }
        table
        {
            border-spacing:0;
            border-collapse: collapse; 
            width:100%;
        }

        table td, table th
        {
            border: 1px solid #ccc;
        }
		
		table th
        {
            background-color:red;
        }
    </style>
</head>
<body>	
    <div class="page">	
        <h1>Table</h1>
        <table border="0">
        <tr>
                <th>No</th>
                <th>Name</th>
                <th>Code</th>
        </tr>
        <?php
        $no = 1;
        foreach($dataProvider->getModels() as $mahasiswa){ 
        ?>
        <tr>
                <td><?= $no++ ?></td>
                <td><?= $mahasiswa->name ?></td>
                <td><?= $mahasiswa->code ?></td>
        </tr>
        <?php
        }
        ?>
        </table>
    </div>   
</body>
</html>
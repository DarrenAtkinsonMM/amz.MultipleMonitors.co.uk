this.Header=function Header()
{
this.SetFont('Arial','',22);
this.SetTextColor(204,204,204);
this.RotatedText(7,47,'Order Invoice',90);
this.SetTextColor(0,0,0);
}
this.Footer=function Footer()
{
this.SetY(-10);
this.SetFont('Arial','',8);
this.Cell(0,10,'Multiple Monitors: ORDER INVOICE ID#2760 (Code: BEP1922507107) - Page '+ this.PageNo()+ '/{nb}',0,0,'L');
}


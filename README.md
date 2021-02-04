# ClipBox-ACAD-NW (proof-of-concept, sample code, use at own risk)
This script only works with a Navisworks Manage Installation and other prerequisites as explained here (see also for usecase):
https://github.com/Henaccount/NW-HideByCoords-Faster see also this video on the usage: https://youtu.be/YZskWnBrZkA

and this video on how the workflow would look like without any script: https://knowledge.autodesk.com/support/autocad-plant-3d/getting-started/caas/simplecontent/content/plant-3d-workflows-using-nwd-coordination-model.html

It has been tested with Navisworks Manage 2020 together with Plant 3D 2021.

To install it, put the resulting dll in a trusted folder of ACAD and load it with "netload".

There will be two commands available: "ClipBoxNw" and "ClipBoxNwReset".
You will use them to manipulate the visible objects in your CoordinationModel.nwd, that you have attached to your active ACAD drawing from the Documents folder.

Usage:
<li>Create a 3D drawing object, that covers partly your Navisworks attachment, e.g. use a box or a sphere. This object will determine the visible area in your NWD.
  <li>Unload the NWD from the xref dialog.
    <li>Call the command "ClipBoxNw" and select that object from the bullet above. Call "ClipBoxNwReset" to show the whole model, if clipped before.
      <li>Now it can take 1 minute or more until the message comes to your command line that the script execution has ended.
        <li>Now reload the NWD from the xref dialog.
          <li>The 3D object can be deleted, but it would make also sense to keep it (on a hidden layer), so you can check what area is save to work in.


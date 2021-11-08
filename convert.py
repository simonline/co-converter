from bs4 import BeautifulSoup
import pandas
import uuid


INPUT_FILENAME = "input.xlsx"
OUTPUT_FILENAME = "output.tt"
BOX_XML = """
<Box>
    <Name>%(name)s</Name>
    <Description></Description>
    <QuantReasoning></QuantReasoning>
    <Color>%(color)i</Color>
    <Attachments/>
    <CLSID>{%(id)s}</CLSID>
    <PosX>0</PosX>
    <PosY>0</PosY>
    <Type>0</Type>
    <State>0</State>
    <IdeaId>0</IdeaId>
    <Position3D>
        <PosX3DResult>0</PosX3DResult>
        <PosY3DResult>0</PosY3DResult>
        <PosZ3DResult>0</PosZ3DResult>
    </Position3D>
</Box>
"""
POINT_XML = """
<OptPoint>
    <CLSID>{%(id)s}</CLSID>
    <X>0</X>
    <Y>0</Y>
    <Z>0</Z>
</OptPoint>
"""
CORRELATION_XML = """
<Correlation>
    <CLSID_ONE>{%(id_1)s}</CLSID_ONE>
    <CLSID_TWO>{%(id_2)s}</CLSID_TWO>
    <Value>%(value)i</Value>
</Correlation>
"""


def convert(input_filename: str, output_filename: str):
    """
    Converts objects and relations from given XLSX file to XML format (with TT extension)

    :param filename:
    :return:
    """
    # Retrieve objects
    df_objects = pandas.read_excel(input_filename, sheet_name="Objekte", header=0)
    objects = [row.values for i, row in df_objects.iterrows()]

    # Retrieve correlations
    df_correlations = pandas.read_excel(input_filename, sheet_name="Korrelationen", header=0)
    correlations = []
    header = list(df_correlations.head(0))
    for i, row in df_correlations.iterrows():
        if i == 0:
            continue
        for j, value in enumerate(row.values):
            if j in (0, 1):
                continue
            if isinstance(value, (float, int)) and not pandas.isna(value):
                correlations.append([int(row.values[0]), int(header[j]), value])

    f = open(output_filename, "r")
    soup = BeautifulSoup(f.read(), "xml")

    # Write boxes and points for objects
    soup.find("Boxes").clear()
    soup.find("OptPointList").clear()
    object_ids = dict(((i[0], uuid.uuid4()) for i in objects))
    for i, object in enumerate(objects):
        box = BeautifulSoup(BOX_XML % {
            "id": object_ids.get(object[0]),
            "name": object[1],
            "color": object[2],
        }, "xml")
        soup.find("Boxes").append(box)

        point = BeautifulSoup(POINT_XML % {
            "id": object_ids.get(object[0]),
        }, "xml")
        soup.find("OptPointList").append(point)

    # Write correlations
    soup.find("Correlations").clear()
    for correlation in correlations:
        box = BeautifulSoup(CORRELATION_XML % {
            "id_1": object_ids.get(correlation[0]),
            "id_2": object_ids.get(correlation[1]),
            "value": correlation[2],
        }, "xml")
        soup.find("Correlations").append(box)

    # Update output file
    f = open(output_filename, "w")
    f.seek(0)
    f.write(soup.prettify(None, "minimal"))
    f.close()


if __name__ == '__main__':
    convert(INPUT_FILENAME, OUTPUT_FILENAME)

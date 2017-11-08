NLA Staff Spreadsheet Importer Plugin
=====================================

An ArchivesSpace (v2.1.x) plugin developed for the National Library of Australia.

It was originally written and maintained by Hudson Molonglo (https://github.com/hudmol/nla_staff_spreadsheet_importer).

After release 2.0, it was moved to and maintained by the NLA.

It adds the following Import Types to the Import Data Job Type.

  * Arrearage spreadsheet
  * Donor Box List spreadsheet
  * Digital Library Collections CSV
  * Basic Resource CSV
  * Obsolete Carriers CSV


Arrearage spreadsheet
---------------------

Takes a spreadsheet
describing a collection (with one record per row) and creates the
following types of records within ArchivesSpace:

  * Resource

  * Archival Object

  * Agent Person

  * Agent Corporate Entity

  * Location

You can find the supported spreadsheet format here:

     https://github.com/nla/nla_staff_spreadsheet_importer/blob/master/samples/Arrearage%20Template.xlsx

The row labelled **ArchivesSpace field code** (marked in grey) is
responsible for mapping each column to a field that the importer knows
about.  Anything you put before this row is ignored.

The row immediately following is the free-text title for each column.
These values aren't used by the importer, so feel free to change them
to labels that you think will be most descriptive for your users.

After the ArchivesSpace field code and title rows, we have one record
per row.  Each record can describe either a Resource record (with a
"Level of description" set to "Collection"), or an Archival Object
record (any other "Level of description" value).  Every row must have
a **Collection Number** value to designate the collection it
defines/belongs to.


Donor Box List spreadsheet
--------------------------

Takes a spreadsheet describing donated collections. A template is provided
to the donor for completion. It is then imported by NLA staff. The template
is included here:

     https://github.com/nla/nla_staff_spreadsheet_importer/blob/master/samples/Donor%20Box%20List%20Template.xlsx

And there are a few examples showing different structures included in the samples directory also:

     https://github.com/nla/nla_staff_spreadsheet_importer/tree/master/samples


Digital Library Collections CSV
-------------------------------

Takes a CSV file exported from the Digital Library Collections system. It creates a single resource
record with class, series, file and item records below it. It also creates linked agent, digital object
and top container records.

A sample CSV file is included here:

     https://github.com/nla/nla_staff_spreadsheet_importer/blob/master/samples/dlc.csv


Basic Resource CSV
------------------

Takes a CSV file and creates a resource record with no components for each row. The resource record
includes rights statement, extent, date, and 'scope and contents' note sub-records.

A sample CSV file is included here:

      https://github.com/nla/nla_staff_spreadsheet_importer/blob/master/samples/basic_resource.csv


Obsolete Carriers CSV
---------------------

Takes a CSV file and creates collection and item level records. The items include instance, extent and
'scope and content' note subrecords, and link to agents and subjects.

A sample CSV files is included here:

      https://github.com/nla/nla_staff_spreadsheet_importer/blob/master/samples/obsolete_carriers.csv


## Installation

### From a released version

Download the latest plugin release from the
[GitHub releases page](https://github.com/nla/nla_staff_spreadsheet_importer/releases).
It will be named something snappy like
`nla_staff_spreadsheet_importer-v0.1.zip`.

Unpack that file into your `/path/to/your/archivesspace/plugins`
directory (yielding
`/path/to/your/archivesspace/plugins/nla_staff_spreadsheet_importer/`).
Next, add the plugin to your ArchivesSpace `config/config.rb` file:

     # If you have other plugins loaded, just add 'nla_staff_spreadsheet_importer' to
     # the list
     AppConfig[:plugins] = ['local', 'nla_staff_spreadsheet_importer']

This plugin needs additional libraries to parse Excel files, so
there's one final step.  From your ArchivesSpace directory, run the
`initialize-plugin` script to install the plugin's dependencies like
this:

     cd /path/to/your/archivesspace
     scripts/initialize-plugin.sh nla_staff_spreadsheet_importer

This will take a minute or two, but you should see it install the
`rubyXL` library.


### From the development version

As above, but instead of downloading and unpacking zip files, just
clone the repository straight into your ArchivesSpace `plugins`
directory like this:

     cd /path/to/your/archivesspace/plugins
     git clone https://github.com/nla/nla_staff_spreadsheet_importer.git nla_staff_spreadsheet_importer


## Configuring it

The Obsolete Carriers CSV importer requires a configuration setting. This will be checked at start up.

     AppConfig[:obsolete_carriers_authorizer_agent_uri] = '/agents/corporate_entities/3'

This should include the uri of an agent record in the system. It will be used as the 'authorizer'
agent for obsolete carrier item records.


## Using it

Once the plugin is loaded, log in to your ArchivesSpace installation
as a user with permission to run imports (the `admin` user is always a
safe bet).  From the  `Create` menu select `Background Jobs`, and this
will load the *New Background Job* page.

From here, select a `Job Type` of `Import Data` and an `Import Type`
of `Arrearage spreadsheet`, `Donor Box List spreadsheet`, or
`Digital Library Collections CSV`.

Finally, click "Add file" and select your spreadsheet file to be
imported or, if your computer has a mouse, you can drag-and-drop the
file straight into ArchivesSpace.

If everything goes well, you should see output that shows records
being created:

     ==================================================
     Arrearage Template.xlsx
     ==================================================
     1. STARTED: Reading JSON records
     1. DONE: Reading JSON records
     2. STARTED: Validating records and checking links
     2. DONE: Validating records and checking links
     3. STARTED: Evaluating record relationships
     3. DONE: Evaluating record relationships
     4. STARTED: Saving records: cycle 1
     Created: /repositories/12345/resources/import_9589708b870be5de95b6ad9f78302f55
     Created: /repositories/12345/resources/import_79b6738635df68deecc401ebbda4a1cd
     Created: /repositories/12345/resources/import_b5fe81985f0368401301aa69f5693f35
     Created: /repositories/12345/resources/import_14cb64e2fec948cbc734fe174fb90225
     Created: /repositories/12345/resources/import_0a6e4f85134a46dac9575c3106ed3dc6
     4. DONE: Saving records: cycle 1
     5. STARTED: Dealing with circular dependencies: cycle 1
     5. DONE: Dealing with circular dependencies: cycle 1
     6. STARTED: Saving records: cycle 2
     Created: /repositories/import/archival_objects/import_bbbefd91-e702-4895-9bd9-4576a58eb937
     Created: /repositories/import/archival_objects/import_bdb6efdf-4dc7-4532-b396-a1fa08d66c1f
     Created: /repositories/import/archival_objects/import_a5788af6-6cec-4375-9c46-54a814badea3
     Created: /repositories/import/archival_objects/import_ccd261c0-7ae2-4c67-a7b0-8ebcbe767207
     6. DONE: Saving records: cycle 2
     7. STARTED: Cleaning up
     7. DONE: Cleaning up


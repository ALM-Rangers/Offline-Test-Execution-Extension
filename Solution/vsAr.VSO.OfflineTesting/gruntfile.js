/// <binding />
//---------------------------------------------------------------------
// <copyright file="gruntfile.js">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>This file in the main entry point for defining grunt tasks and using grunt plugins.
// Click here to learn more. http://go.microsoft.com/fwlink/?LinkID=513275&clcid=0x409
// </summary>
//---------------------------------------------------------------------

module.exports = function (grunt) {
    grunt.initConfig({
        pkg: grunt.file.readJSON("package.json"),
        exec: {
            package: {
                command: "tfx extension create --manifest-globs vss-extension.json",
                stdout: true,
                stderr: true
            },
            install: {
                command: "npm install --save vss-web-extension-sdk",
                stdout: true,
                stderr: true
            },
            
            update: {
                command: "npm up --save",
                stdout: true,
                stderr: true
            },
            publish: {
                command: "tfx extension publish --token ewjvmszxduyb4r7wvjurydihvkrj37ptaz45f3jgpuqx7gaj5nua",
                stdout: true,
                stderr: true
            },
            publish_local: {
                command: "tfx extension publish --root . --manifest-globs extension.onprem.json --service-url http://localhost:8080/tfs",
                stdout: true,
                stderr: true
            }
        },
        copy: {
            main: {
                files: [
                    // includes files within path
                    { expand: true, flatten: true, src: ['node_modules/vss-web-extension-sdk/lib/VSS.SDK.js'], dest: 'scripts/', filter: 'isFile' }]
            }
        },
        jasmine: {
            src: ["scripts/**/*.js", "sdk/scripts/*.js"],
            specs: "test/**/*[sS]pec.js",
            helpers: "test/helpers/*.js"
        },
        typescript: {
            base: {
                src: ['scripts/**/*.ts'],
                options: {
                    module: 'amd', //or commonjs 
                    moduleResolution: "node",
                }
            }
        }
    });

    grunt.loadNpmTasks("grunt-exec");
    grunt.loadNpmTasks("grunt-contrib-copy");
    grunt.loadNpmTasks("grunt-contrib-jasmine");
    grunt.loadNpmTasks("grunt-typescript");
    grunt.registerTask("vscode", ["exec:update", "copy:main", "typescript", "exec:package"]);
};
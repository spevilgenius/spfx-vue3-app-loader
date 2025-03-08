/* eslint-disable @typescript-eslint/no-require-imports */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
const fs = require('fs')
const process = require('process')

const VERSION_TYPE = process.argv[2]
const PACKAGE_SOLUTION = '/config/package-solution.json'
// const WEBPART_BASE = '/src/webparts'

const getPackageVersion = (basePath) => {
  try {
    let pkg = JSON.parse(fs.readFileSync(basePath + 'package.json', 'utf-8'))
    return pkg.version
  } catch (e) {
    console.log('ERROR: ' + e)
    return ''
  }
}

const updatePackageVersion = (basePath) => {
  try {
    let ps = JSON.parse(fs.readFileSync(basePath + PACKAGE_SOLUTION, 'utf-8'))
    let version = String(ps.solution.version)
    let newversion = ''
    let tp1 = version.split('.')
    switch(VERSION_TYPE) {
      case 'Major': {
        let tp2 = Number(tp1[0])
        tp2 += 1
        newversion = String(tp2 + '.0.0.0')
        break
      }
      case 'Minor': {
        let tp2 = Number(tp1[1])
        tp2 += 1
        newversion = String(tp1[0] + '.' + tp2 + '.0.0')
        break
      }
      case 'Draft': {
        let tp2 = Number(tp1[2])
        tp2 += 1
        newversion = String(tp1[0] + '.' + tp1[1] + '.' + tp2 + '.0')
        break
      }
    }
    ps.solution.version = newversion
    fs.writeFileSync(basePath + PACKAGE_SOLUTION, JSON.stringify(ps, null, '\t'))
    console.log('Version updated to ' + newversion)
  } catch (e) {
    // don't care
  }
}

run = () => {
  let basePath = process.cwd()  + '\\'
  basePath += '\\'
  let version = getPackageVersion(basePath)
  basePath = process.cwd()  + '\\'
  if (version) {
    updatePackageVersion(basePath, version + '.0')
  }
}

run()